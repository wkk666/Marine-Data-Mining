"""模拟海区的最小 H3 查询实验；不使用真实海区或真实 AIS。

python run_experiment.py --prepare-repeats 7 --query-repeats 31
结果写入独立 runs/时间戳 目录。正确性失败时保存诊断并停止计时。
"""
from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from datetime import datetime
import gc
import hashlib
import json
import os
from pathlib import Path
import platform
import random
import statistics
import sys
import time

ROOT = Path(__file__).resolve().parent
LOCAL_PACKAGES = ROOT / ".runtime" / "packages"
if LOCAL_PACKAGES.is_dir():
    sys.path.insert(0, str(LOCAL_PACKAGES))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import numpy as np
import shapely
from h3.api import basic_int as h3

SEED = 20261110
RESOLUTIONS = (6, 7, 8)
METHODS = ("direct", "h3_r6", "h3_r7", "h3_r8")
# 二进制可精确表示的顶点与边中点，避免“恰在边界”被十进制插值误差改变。
VERTICES = ((122.0, 30.0), (122.25, 30.0), (122.25, 30.25),
            (122.15625, 30.25), (122.15625, 30.09375),
            (122.09375, 30.09375), (122.09375, 30.25), (122.0, 30.25))
EXTENT = (121.875, 29.875, 122.375, 30.375)
EDGE_OFFSET_DEG = 1e-7
EMPTY_ROWS = np.empty(0, dtype=np.int64)


@dataclass
class Dataset:
    ids: np.ndarray
    lon: np.ndarray
    lat: np.ndarray
    category: list[str]
    expected_probe_membership: list[bool | None]


@dataclass
class State:
    name: str
    region: object
    bounds: tuple
    index: dict | None = None
    encoded: np.ndarray | None = None
    enumerated_cells: tuple = ()
    cover: tuple = ()
    internal: tuple = ()
    border: tuple = ()
    cell_geometries: object = None


def dump_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_csv(path: Path, rows: list[dict], fieldnames=None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames or list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def generate_data(n=10000, seed=SEED) -> Dataset:
    rng = np.random.Generator(np.random.PCG64(seed))
    lon = rng.uniform(EXTENT[0], EXTENT[2], n).tolist()
    lat = rng.uniform(EXTENT[1], EXTENT[3], n).tolist()
    category = ["uniform_simulated"] * n
    expected = [None] * n
    # CCW 多边形每条边的左侧为内部。每边增加中点、内部近邻、外部近邻。
    for edge, (start, end) in enumerate(zip(VERTICES, VERTICES[1:] + VERTICES[:1])):
        a, b = np.asarray(start), np.asarray(end)
        midpoint = (a + b) / 2
        delta = b - a
        inward = np.array([-delta[1], delta[0]]) / np.linalg.norm(delta)
        for label, sign, truth in (("on_edge", 0, True), ("near_inside", 1, True),
                                    ("near_outside", -1, False)):
            point = midpoint + sign * EDGE_OFFSET_DEG * inward
            lon.append(float(point[0]))
            lat.append(float(point[1]))
            category.append(f"{label}_{edge}")
            expected.append(truth)
    for i, (x, y) in enumerate(VERTICES):
        lon.append(x)
        lat.append(y)
        category.append(f"on_vertex_{i}")
        expected.append(True)
    return Dataset(np.arange(1, len(lon) + 1, dtype=np.int64),
                   np.asarray(lon), np.asarray(lat), category, expected)


def make_region():
    region = shapely.Polygon(VERTICES)
    shapely.prepare(region)
    # GEOS 的部分预处理可能延迟到首次谓词。这里显式触发，费用归入准备。
    shapely.intersects_xy(region, np.array([122.03125]), np.array([30.03125]))
    return region


def build_point_index(data: Dataset, resolution: int):
    start = time.perf_counter_ns()
    # 每一档独立从原始纬经度编码，不由其他档的父子索引推算。
    encoded = np.fromiter((h3.latlng_to_cell(float(lat), float(lon), resolution)
                           for lon, lat in zip(data.lon, data.lat)),
                          dtype=np.uint64, count=len(data.ids))
    after_encoding = time.perf_counter_ns()
    order = np.argsort(encoded, kind="stable")
    unique, starts = np.unique(encoded[order], return_index=True)
    ends = np.r_[starts[1:], len(order)]
    index = {int(cell): order[a:b] for cell, a, b in zip(unique, starts, ends)}
    end = time.perf_counter_ns()
    return encoded, index, after_encoding - start, end - after_encoding


def enumerate_cover(resolution: int, vertices=VERTICES) -> tuple[int, ...]:
    shape = h3.LatLngPoly([(lat, lon) for lon, lat in vertices])
    # bbox_overlap 是保守枚举，不使用可能遗漏边缘的中心落入模式。
    return tuple(sorted(h3.h3shape_to_cells_experimental(shape, resolution,
                                                         contain="bbox_overlap")))


def classify_cells(region, enumerated):
    polygons = np.asarray([shapely.Polygon([(lon, lat) for lat, lon in h3.cell_to_boundary(cell)])
                           for cell in enumerated], dtype=object)
    # intersects 包含仅接触边界的格子；covers 含边界，不使用中心点替代。
    intersects = np.asarray(shapely.intersects(region, polygons), dtype=bool)
    full = np.asarray(shapely.covers(region, polygons), dtype=bool)
    cells = np.asarray(enumerated, dtype=np.uint64)
    cover = tuple(map(int, cells[intersects]))
    internal = tuple(map(int, cells[intersects & full]))
    border = tuple(map(int, cells[intersects & ~full]))
    return cover, internal, border, polygons[intersects]


def prepare_method(name: str, data: Dataset):
    start = time.perf_counter_ns()
    region = make_region()
    bounds = region.bounds
    after_geometry = time.perf_counter_ns()
    state = State(name, region, bounds)
    metrics = {"geometry_prepare_ns": after_geometry - start, "point_index_ns": 0,
               "point_encoding_ns": 0, "index_grouping_ns": 0,
               "cover_generation_ns": 0, "cell_classification_ns": 0}
    if name != "direct":
        resolution = int(name[-1])
        index_start = time.perf_counter_ns()
        state.encoded, state.index, encode_ns, grouping_ns = build_point_index(data, resolution)
        index_end = time.perf_counter_ns()
        state.enumerated_cells = enumerate_cover(resolution)
        cover_end = time.perf_counter_ns()
        state.cover, state.internal, state.border, state.cell_geometries = classify_cells(
            region, state.enumerated_cells)
        classify_end = time.perf_counter_ns()
        metrics.update(point_index_ns=index_end - index_start, point_encoding_ns=encode_ns,
                       index_grouping_ns=grouping_ns, cover_generation_ns=cover_end - index_end,
                       cell_classification_ns=classify_end - cover_end)
    end = time.perf_counter_ns()
    metrics["prepare_total_ns"] = end - start
    metrics["prepare_other_ns"] = metrics["prepare_total_ns"] - sum(
        metrics[k] for k in ("geometry_prepare_ns", "point_index_ns", "cover_generation_ns",
                              "cell_classification_ns"))
    return state, metrics


def concatenate_rows(chunks):
    return np.concatenate(chunks) if chunks else EMPTY_ROWS


def query(state: State, data: Dataset):
    """相同输出契约：返回唯一记录 ID 的一维数组，不保证顺序。

    所有候选提取、坐标收集、精判、结果数组构造都计时；不含集合比较和写盘。
    """
    start = time.perf_counter_ns()
    if state.name == "direct":
        xmin, ymin, xmax, ymax = state.bounds
        rows = np.flatnonzero((data.lon >= xmin) & (data.lon <= xmax)
                              & (data.lat >= ymin) & (data.lat <= ymax))
        internal_rows = EMPTY_ROWS
    else:
        internal_chunks = []
        border_chunks = []
        for cell in state.internal:
            group = state.index.get(cell)
            if group is not None:
                internal_chunks.append(group)
        for cell in state.border:
            group = state.index.get(cell)
            if group is not None:
                border_chunks.append(group)
        internal_rows = concatenate_rows(internal_chunks)
        rows = concatenate_rows(border_chunks)
    x, y = data.lon[rows], data.lat[rows]
    exact_start = time.perf_counter_ns()
    # 对点而言 intersects_xy 与 covers(region, Point) 同样包含边界，免建 Point 对象。
    selected = shapely.intersects_xy(state.region, x, y)
    exact_end = time.perf_counter_ns()
    accepted = rows[selected]
    if state.name != "direct":
        accepted = np.concatenate((internal_rows, accepted))
    ids = data.ids[accepted]
    end = time.perf_counter_ns()
    return ids, {"query_total_ns": end - start, "lookup_and_gather_ns": exact_start - start,
                 "exact_predicate_ns": exact_end - exact_start, "assemble_result_ns": end - exact_end,
                 "candidate_points": len(rows) + len(internal_rows), "exact_points": len(rows),
                 "internal_accepted_points": len(internal_rows), "selected_points": len(ids)}


def check_coverage(state: State) -> dict:
    resolution = int(state.name[-1])
    # 几何并集审计检查整片海区，不只检查已生成的数据点。
    union = shapely.union_all(state.cell_geometries)
    gap = shapely.difference(state.region, union)
    union_covers = bool(shapely.covers(union, state.region))
    xmin, ymin, xmax, ymax = state.bounds
    # 另从更大的矩形枚举，核对凹口形状的枚举没有漏掉相交格。
    margin = 0.0625
    rectangle = ((xmin - margin, ymin - margin), (xmax + margin, ymin - margin),
                 (xmax + margin, ymax + margin), (xmin - margin, ymax + margin))
    broad = enumerate_cover(resolution, rectangle)
    broad_cover, _, _, _ = classify_cells(state.region, broad)
    missing = sorted(set(broad_cover) - set(state.cover))
    unexpected = sorted(set(state.cover) - set(broad_cover))
    return {"enumeration_mode": "bbox_overlap_then_planar_intersects",
            "enumerated_cells": len(state.enumerated_cells), "cover_cells": len(state.cover),
            "internal_cells": len(state.internal), "border_cells": len(state.border),
            "discarded_nonintersecting_envelope_cells": len(state.enumerated_cells) - len(state.cover),
            "union_covers_entire_region": union_covers, "uncovered_area_square_degrees": float(gap.area),
            "extended_rectangle_cells": len(broad), "missing_vs_extended_rectangle": missing,
            "extra_vs_extended_rectangle": unexpected,
            "passed": union_covers and not missing and not unexpected}


def save_dataset(run: Path, data: Dataset, seed: int, n: int) -> None:
    rows = [{"record_id": int(record_id), "longitude": format(lon, ".17g"),
             "latitude": format(lat, ".17g"), "category": category,
             "probe_expected_inside": "" if truth is None else int(truth)}
            for record_id, lon, lat, category, truth in zip(data.ids, data.lon, data.lat,
                                                         data.category, data.expected_probe_membership)]
    write_csv(run / "data" / "simulated_points.csv", rows)
    dump_json(run / "data" / "simulated_sea.geojson", {
        "type": "Feature", "properties": {"simulated": True, "name": "SIMULATED_NOTCHED_SEA"},
        "geometry": shapely.Polygon(VERTICES).__geo_interface__})
    dump_json(run / "data" / "generation_rules.json", {
        "all_data_simulated": True, "seed": seed, "rng": "numpy.Generator(PCG64)",
        "uniform_points": n, "probe_points": len(data.ids) - n, "total_points": len(data.ids),
        "longitude_latitude_extent": EXTENT, "polygon_vertices_lon_lat": VERTICES,
        "uniform_distribution": "independent uniform longitude and latitude; not equal-area uniform",
        "probe_rule": "each of 8 edges: exact midpoint + inward/outward offset; also all 8 vertices",
        "probe_offset_degrees": EDGE_OFFSET_DEG, "boundary_included": True,
        "geometry_model": "local planar straight segments in longitude/latitude; no geodesic-distance claims",
        "records": "unique observation IDs; one fixed synthetic snapshot; no AIS or vessel deduplication"})


def correctness_phase(run: Path, data: Dataset):
    states = {}
    summaries = {}
    reference_region = shapely.Polygon(VERTICES)
    # 验证用 covers + Point 对象；计时用更快的 intersects_xy，不将验证成本计入查询。
    oracle_mask = np.asarray(shapely.covers(reference_region, shapely.points(data.lon, data.lat)))
    oracle = set(map(int, data.ids[oracle_mask]))
    results = {}
    diagnostics = []
    cover_ok = True
    for method in METHODS:
        state, _ = prepare_method(method, data)
        states[method] = state
        selected, counts = query(state, data)
        actual = set(map(int, selected))
        false_positive = actual - oracle
        false_negative = oracle - actual
        results[method] = actual
        duplicate_count = len(selected) - len(actual)
        coverage = None if method == "direct" else check_coverage(state)
        if coverage is not None:
            cover_ok = cover_ok and coverage["passed"]
        summaries[method] = {"false_positive": len(false_positive), "false_negative": len(false_negative),
                             "duplicate_results": duplicate_count, **{k: v for k, v in counts.items()
                                                                        if not k.endswith("_ns")},
                             "coverage": coverage}
        for row, record_id in enumerate(data.ids):
            if int(record_id) in false_positive | false_negative:
                encoded = "" if state.encoded is None else h3.int_to_str(int(state.encoded[row]))
                diagnostics.append({"method": method, "record_id": int(record_id),
                                    "longitude": data.lon[row], "latitude": data.lat[row],
                                    "category": data.category[row], "h3_cell": encoded,
                                    "expected_inside": int(oracle_mask[row]),
                                    "actual_inside": int(int(record_id) in actual)})
        write_csv(run / "selected_ids" / f"{method}.csv",
                  [{"record_id": record_id} for record_id in sorted(actual)], ["record_id"])
    edge_rows = []
    for row, expected in enumerate(data.expected_probe_membership):
        if expected is None:
            continue
        record_id = int(data.ids[row])
        result = {"record_id": record_id, "category": data.category[row],
                  "longitude": format(data.lon[row], ".17g"), "latitude": format(data.lat[row], ".17g"),
                  "expected_inside": int(expected), "oracle_inside": int(oracle_mask[row]),
                  "exactly_on_boundary": int(shapely.intersects_xy(shapely.boundary(reference_region),
                                                                    data.lon[row], data.lat[row]))}
        result.update({method: int(record_id in results[method]) for method in METHODS})
        for method in METHODS[1:]:
            state = states[method]
            cell = int(state.encoded[row])
            result[f"{method}_cell"] = h3.int_to_str(cell)
            result[f"{method}_class"] = ("internal" if cell in state.internal else
                                          "border" if cell in state.border else "outside_cover")
        result["passed"] = all(bool(result[method]) == expected for method in METHODS) and bool(
            oracle_mask[row]) == expected
        edge_rows.append(result)
    probe_ok = all(row["passed"] for row in edge_rows)
    passed = not diagnostics and cover_ok and probe_ok and all(
        s["duplicate_results"] == 0 for s in summaries.values())
    checks = {"passed": passed, "oracle_selected_points": len(oracle), "total_points": len(data.ids),
              "probe_points": len(edge_rows), "probe_checks_passed": probe_ok,
              "all_cover_checks_passed": cover_ok, "methods": summaries}
    dump_json(run / "correctness.json", checks)
    write_csv(run / "boundary_cases.csv", edge_rows)
    write_csv(run / "diagnostics" / "incorrect_points.csv", diagnostics,
              ["method", "record_id", "longitude", "latitude", "category", "h3_cell",
               "expected_inside", "actual_inside"])
    # 先保存诊断再停止；不会对错误算法继续做性能结论。
    if not passed:
        dump_json(run / "status.json", {"status": "correctness_failed", "performance_run": False})
        raise RuntimeError(f"正确性检查失败，已保存出错点和边界检查：{run}")
    for method, state in states.items():
        if method == "direct":
            continue
        cover_rows = [{"h3_cell": h3.int_to_str(cell),
                       "class": "internal" if cell in state.internal else "border"}
                      for cell in state.cover]
        write_csv(run / "coverage" / f"{method}_cells.csv", cover_rows)
        write_csv(run / "point_indexes" / f"{method}_encoding.csv",
                  [{"record_id": int(record_id), "h3_cell": h3.int_to_str(int(cell))}
                   for record_id, cell in zip(data.ids, state.encoded)])
    return checks, oracle


def benchmark(data: Dataset, oracle: set[int], prepare_repeats: int, query_repeats: int, seed: int):
    randomizer = random.Random(seed + 1)
    prepare_trials, query_trials = [], []
    gc_was_enabled = gc.isenabled()
    gc.disable()
    try:
        # 各方法预热一次，导入、动态链接、首次分配不混进核心计时。
        for method in METHODS:
            warm_state, _ = prepare_method(method, data)
            query(warm_state, data)
        del warm_state
        for round_id in range(1, prepare_repeats + 1):
            order = list(METHODS)
            randomizer.shuffle(order)
            for position, method in enumerate(order, 1):
                state, preparation = prepare_method(method, data)
                ids, first_query = query(state, data)
                assert set(map(int, ids)) == oracle
                prepare_trials.append({"trial": round_id, "order_in_round": position,
                                       "method": method, **preparation,
                                       **{f"first_{k}": v for k, v in first_query.items()},
                                       "one_shot_total_ns": preparation["prepare_total_ns"]
                                                            + first_query["query_total_ns"]})
        states = {method: prepare_method(method, data)[0] for method in METHODS}
        for state in states.values():
            for _ in range(3):
                query(state, data)
        for round_id in range(1, query_repeats + 1):
            order = list(METHODS)
            randomizer.shuffle(order)
            for position, method in enumerate(order, 1):
                ids, metrics = query(states[method], data)
                # ID 集合构造和校验始终放在 query 内部计时之外。
                assert set(map(int, ids)) == oracle
                query_trials.append({"trial": round_id, "order_in_round": position,
                                     "method": method, **metrics})
        return prepare_trials, query_trials
    finally:
        if gc_was_enabled:
            gc.enable()


def summarize(preparation: list[dict], queries: list[dict], checks: dict) -> list[dict]:
    rows = []
    for method in METHODS:
        p = [r for r in preparation if r["method"] == method]
        q = [r for r in queries if r["method"] == method]
        correctness = checks["methods"][method]
        cover = correctness["coverage"]
        med = lambda rs, key: statistics.median(row[key] for row in rs) / 1e6
        row = {"method": method, "prepare_trials": len(p), "query_trials": len(q),
               "selected_points": correctness["selected_points"],
               "false_positive": correctness["false_positive"], "false_negative": correctness["false_negative"],
               "cover_cells": 0 if cover is None else cover["cover_cells"],
               "internal_cells": 0 if cover is None else cover["internal_cells"],
               "border_cells": 0 if cover is None else cover["border_cells"],
               "candidate_points": correctness["candidate_points"], "exact_points": correctness["exact_points"],
               "geometry_prepare_ms": med(p, "geometry_prepare_ns"),
               "point_index_ms": med(p, "point_index_ns"), "point_encoding_ms": med(p, "point_encoding_ns"),
               "index_grouping_ms": med(p, "index_grouping_ns"),
               "cover_generation_ms": med(p, "cover_generation_ns"),
               "cell_classification_ms": med(p, "cell_classification_ns"),
               "prepare_total_ms": med(p, "prepare_total_ns"),
               "first_query_ms": med(p, "first_query_total_ns"),
               "one_shot_total_ms": med(p, "one_shot_total_ns"),
               "query_ms": med(q, "query_total_ns"), "exact_predicate_ms": med(q, "exact_predicate_ns"),
               "lookup_and_gather_ms": med(q, "lookup_and_gather_ns"),
               "assemble_result_ms": med(q, "assemble_result_ns"),
               "query_min_ms": min(r["query_total_ns"] for r in q) / 1e6,
               "query_max_ms": max(r["query_total_ns"] for r in q) / 1e6}
        rows.append(row)
    return rows


def make_quick_results(run: Path, summaries: list[dict], checks: dict, args):
    lines = ["# 模拟海区查询：本轮结果速览", "", "海区和全部船位均为模拟数据，不是真实海区或 AIS。", "",
             f"原始记录、数据和完整参数：[本次运行]({run.relative_to(ROOT).as_posix()}/summary.json)。", "",
             f"共 {checks['total_points']} 个点（{args.points:,} 个均匀随机点 + 32 个边界测试点），"
             f"最终选中 {checks['oracle_selected_points']} 个。", "",
             "所有时间单位为毫秒；各列为各自重复记录的中位数，不将分项中位数相加。", "",
             "| 方案 | 点索引 | 覆盖枚举 | 格子分类 | 准备总计 | 准备后查询 | 其中精判 | 一次性准备+首查 |",
             "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |"]
    for r in summaries:
        lines.append(f"| {r['method']} | {r['point_index_ms']:.4f} | {r['cover_generation_ms']:.4f} | "
                     f"{r['cell_classification_ms']:.4f} | {r['prepare_total_ms']:.4f} | {r['query_ms']:.4f} | "
                     f"{r['exact_predicate_ms']:.4f} | {r['one_shot_total_ms']:.4f} |")
    lines += ["", "直接方法的“精判”针对外包矩形初筛后的全部候选点；H3 针对边界格点。", "",
              "| 方案 | 覆盖格 | 内部格 | 边界格 | 候选点 | 实际精判点 | 误选 | 漏选 |",
              "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |"]
    for r in summaries:
        lines.append(f"| {r['method']} | {r['cover_cells']} | {r['internal_cells']} | {r['border_cells']} | "
                     f"{r['candidate_points']} | {r['exact_points']} | {r['false_positive']} | {r['false_negative']} |")
    lines += ["", "直接方案不使用 H3 格子，因此格子数记为 0。", "",
              "## 计时与正确性口径", "",
              "- 直接方案：NumPy 外包矩形初筛 + 已预处理几何上的 Shapely 向量化 intersects_xy；海区边界计入。",
              "- H3：各档从原始经纬度编码并按编号分组；每次重新查格、收集候选、边界精判、组装 ID，未缓存查询答案。",
              "- 覆盖：bbox_overlap 保守枚举，再以格子实际轮廓与海区 intersects 过滤；covers 划分内部格，其余相交格为边界格。",
              "- 审计：覆盖格并集覆盖整片海区，并与扩大矩形枚举得到的全部相交格交叉核对；正确性比较唯一记录 ID 集合。",
              "- 独立预期检查：8 个边中点、8 个顶点应纳入；8 个内侧近邻纳入；8 个外侧近邻排除。",
              f"- 准备计时重复 {args.prepare_repeats} 次，准备后查询重复 {args.query_repeats} 次。每轮随机打乱方案顺序。各方法均预热，GC 统一暂停。",
              "- 数据生成、文件读写、正确性检查、集合比较、覆盖审计、依赖导入不计入核心时间。",
              "- 一次性总成本为同一试次的准备时间与首查时间之和，再取中位数；不含进程启动。",
              "- 数据是同一快照；重复查询只复用索引与海区清单。数据/海区更新后的重建成本未模拟。", "",
              "这是一个小规模、单海区、均匀模拟点和特定 Python 实现的观察，尚未设计分辨率选择器。", ""]
    content = "\n".join(lines)
    (ROOT / "结果速览.md").write_text(content, encoding="utf-8")
    (run / "结果速览.md").write_text(content.replace(f"{run.relative_to(ROOT).as_posix()}/summary.json",
                                                     "summary.json"), encoding="utf-8")


def environment_info():
    cpu = platform.processor()
    if sys.platform == "win32":
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DESCRIPTION\System\CentralProcessor\0") as key:
            cpu = winreg.QueryValueEx(key, "ProcessorNameString")[0].strip()
    return {"timestamp": datetime.now().astimezone().isoformat(), "python": sys.version,
            "executable": sys.executable, "platform": platform.platform(), "cpu": cpu,
            "logical_cpus": os.cpu_count(), "numpy": np.__version__, "shapely": shapely.__version__,
            "geos": shapely.geos_version_string, "h3": h3.versions(),
            "clock": str(time.get_clock_info("perf_counter")), "single_process": True,
            "method_order_seed": SEED + 1, "gc_disabled_during_benchmark": True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--seed", type=int, default=SEED)
    parser.add_argument("--points", type=int, default=10000)
    parser.add_argument("--prepare-repeats", type=int, default=7)
    parser.add_argument("--query-repeats", type=int, default=31)
    args = parser.parse_args()
    if min(args.prepare_repeats, args.query_repeats) < 5:
        parser.error("每类计时至少重复 5 次")
    if args.points < 1:
        parser.error("随机点数量须大于 0")
    run = ROOT / "runs" / datetime.now().astimezone().strftime("%Y%m%d_%H%M%S_%f")
    run.mkdir(parents=True)
    print(f"模拟实验输出：{run}", flush=True)
    data = generate_data(args.points, args.seed)
    assert len(set(map(int, data.ids))) == len(data.ids)
    save_dataset(run, data, args.seed, args.points)
    env = environment_info()
    env["method_order_seed"] = args.seed + 1
    dump_json(run / "environment.json", env)
    dump_json(run / "parameters.json", vars(args))
    (run / "run_experiment_snapshot.py").write_bytes(Path(__file__).read_bytes())
    (run / "requirements_snapshot.txt").write_bytes((ROOT / "requirements.txt").read_bytes())
    print("先检查覆盖完整性、记录 ID、边界测试点。", flush=True)
    checks, oracle = correctness_phase(run, data)
    print(f"正确性通过：4 个方案均选中 {len(oracle)} 个点，32 个边界测试点全部通过。", flush=True)
    print("开始分别计时准备工作、首查、准备后重复查询。", flush=True)
    preparations, queries = benchmark(data, oracle, args.prepare_repeats, args.query_repeats, args.seed)
    write_csv(run / "preparation_trials.csv", preparations)
    write_csv(run / "query_trials.csv", queries)
    summaries = summarize(preparations, queries, checks)
    write_csv(run / "summary.csv", summaries)
    dump_json(run / "summary.json", {"all_data_simulated": True, "parameters": vars(args),
                                     "correctness": checks, "results": summaries})
    make_quick_results(run, summaries, checks, args)
    dump_json(run / "status.json", {"status": "completed", "correctness_passed": True,
                                    "performance_trials": len(queries), "preparation_trials": len(preparations)})
    manifest = {p.relative_to(run).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
                for p in run.rglob("*") if p.is_file()}
    manifest["../../run_experiment.py"] = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    dump_json(run / "sha256_manifest.json", manifest)
    (ROOT / "LATEST_RUN.txt").write_text(run.relative_to(ROOT).as_posix() + "\n", encoding="utf-8")
    for r in summaries:
        print(f"{r['method']}: 准备 {r['prepare_total_ms']:.4f} ms，重复查询 {r['query_ms']:.4f} ms，"
              f"精判 {r['exact_predicate_ms']:.4f} ms，误选/漏选 {r['false_positive']}/{r['false_negative']}")
    print("已保存每次原始计时、模拟数据、边界检查及结果。", flush=True)


if __name__ == "__main__":
    main()
