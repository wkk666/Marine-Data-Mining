"""复用四组已保存数据，仅增加H3边界候选的含边界外包矩形筛选。"""
from datetime import datetime
import csv
import gc
import hashlib
import json
from pathlib import Path
import random
import shutil
import statistics
import time

import run_experiment as base
from exploration_experiment import distribution_stats

np, shapely, h3 = base.np, base.shapely, base.h3
ROOT = Path(__file__).resolve().parent
PROTOCOL_PATH = ROOT / "bbox_ablation_protocol.json"


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream))


def load_saved_data(folder):
    saved_region = json.loads((folder / "data" / "simulated_sea.geojson").read_text("utf-8"))
    region = shapely.Polygon(saved_region["geometry"]["coordinates"][0])
    assert shapely.equals_exact(region, shapely.Polygon(base.VERTICES), 0)
    rows = read_csv(folder / "data" / "simulated_points.csv")
    data = base.Dataset(np.array([int(r["record_id"]) for r in rows], dtype=np.int64),
                        np.array([float(r["longitude"]) for r in rows]),
                        np.array([float(r["latitude"]) for r in rows]),
                        [r["category"] for r in rows],
                        [None if r["probe_expected_inside"] == "" else bool(int(r["probe_expected_inside"]))
                         for r in rows])
    assert len(data.ids) == len(set(map(int, data.ids))) == 10032
    for array in (data.ids, data.lon, data.lat):
        array.flags.writeable = False
    return data


def query_h3_bbox(state, data, profile=False):
    """H3原流程+边界候选bbox；主计时不启用过滤专用探针。

    normal metrics 与原查询的结构相同。profile=True仅用于独立细分诊断。
    """
    start = time.perf_counter_ns()
    internal_chunks, border_chunks = [], []
    for cell in state.internal:
        group = state.index.get(cell)
        if group is not None:
            internal_chunks.append(group)
    for cell in state.border:
        group = state.index.get(cell)
        if group is not None:
            border_chunks.append(group)
    internal_rows = base.concatenate_rows(internal_chunks)
    rows = base.concatenate_rows(border_chunks)
    x, y = data.lon[rows], data.lat[rows]
    if profile:
        bbox_start = time.perf_counter_ns()
    xmin, ymin, xmax, ymax = state.bounds
    keep = (x >= xmin) & (x <= xmax) & (y >= ymin) & (y <= ymax)
    if profile:
        after_mask = time.perf_counter_ns()
    # 初始坐标提取、mask、三数组筛选均在查询内部；不使用预先筛好的行号。
    rows, x, y = rows[keep], x[keep], y[keep]
    if profile:
        after_compaction = time.perf_counter_ns()
    exact_start = time.perf_counter_ns()
    selected = shapely.intersects_xy(state.region, x, y)
    exact_end = time.perf_counter_ns()
    accepted = rows[selected]
    accepted = np.concatenate((internal_rows, accepted))
    ids = data.ids[accepted]
    end = time.perf_counter_ns()
    metrics = {"query_total_ns": end - start, "lookup_and_gather_ns": exact_start - start,
               "exact_predicate_ns": exact_end - exact_start, "assemble_result_ns": end - exact_end,
               "candidate_points": len(rows) + len(internal_rows), "exact_points": len(rows),
               "internal_accepted_points": len(internal_rows), "selected_points": len(ids)}
    if profile:
        metrics.update(bbox_mask_ns=after_mask - bbox_start,
                       bbox_array_compaction_ns=after_compaction - after_mask,
                       bbox_filter_ns=after_compaction - bbox_start,
                       pre_bbox_gather_ns=bbox_start - start,
                       border_candidates_before_bbox=len(keep),
                       bbox_removed_points=len(keep) - len(rows))
    return ids, metrics


def original_method(method):
    return method.removesuffix("_bbox")


def run_query(method, state, data, profile=False):
    if method.endswith("_bbox"):
        return query_h3_bbox(state, data, profile=profile)
    return base.query(state, data)


def preflight(run, source, datasets, protocol):
    checks, oracles, counts = {}, {}, []
    for name, data in datasets.items():
        folder = run / "datasets" / name
        check, oracle = base.correctness_phase(folder, data)
        variants = {}
        for resolution in (6, 7, 8):
            method = f"h3_r{resolution}"
            state, _ = base.prepare_method(method, data)
            border_rows = base.concatenate_rows([state.index[cell] for cell in state.border if cell in state.index])
            x, y = data.lon[border_rows], data.lat[border_rows]
            xmin, ymin, xmax, ymax = state.bounds
            within = (x >= xmin) & (x <= xmax) & (y >= ymin) & (y <= ymax)
            for use_profile in (False, True):
                selected, metrics = query_h3_bbox(state, data, profile=use_profile)
                actual = set(map(int, selected))
                extra, missing = actual - oracle, oracle - actual
                duplicate = len(selected) - len(actual)
                if extra or missing or duplicate:
                    base.dump_json(folder / "bbox_failure.json", {"method": method, "profile": use_profile,
                                   "false_positive_ids": sorted(extra), "false_negative_ids": sorted(missing),
                                   "duplicate_ids": duplicate})
                    raise RuntimeError("bbox变体正确性失败；已保存诊断，不做性能结论")
            probes = [i for i, expected in enumerate(data.expected_probe_membership) if expected is not None]
            assert all((int(data.ids[i]) in actual) == data.expected_probe_membership[i] for i in probes)
            variants[method + "_bbox"] = {"false_positive": 0, "false_negative": 0, "duplicate_results": 0,
                                            "probe_checks_passed": True, "selected_points": len(selected)}
            base.write_csv(folder / "selected_ids" / f"{method}_bbox.csv",
                           [{"record_id": i} for i in sorted(actual)], ["record_id"])
            # 与上一轮保存的编码及相交格清单交叉核对，确认本轮确实是同一输入。
            previous_enc = read_csv(source / "datasets" / name / "point_indexes" / f"{method}_encoding.csv")
            assert all(int(row["record_id"]) == int(record_id) and int(row["h3_cell"], 16) == int(cell)
                       for row, record_id, cell in zip(previous_enc, data.ids, state.encoded, strict=True))
            old_ids = {int(row["record_id"]) for row in read_csv(source / "datasets" / name / "selected_ids" / f"{method}.csv")}
            assert old_ids == oracle
            counts.append({"dataset": name, "resolution": resolution,
                           "border_candidates": len(border_rows), "bbox_removed": int((~within).sum()),
                           "new_exact_points": int(within.sum()),
                           "original_exact_points": check["methods"][method]["exact_points"],
                           "internal_accepted_points": metrics["internal_accepted_points"],
                           "selected_points": len(selected)})
            base.write_csv(folder / f"{method}_bbox_removed_ids.csv",
                           [{"record_id": int(i)} for i in data.ids[border_rows[~within]]], ["record_id"])
        checks[name] = {"original": check, "bbox_variants": variants}
        oracles[name] = oracle
        print(f"{name}：7方案ID集合相同，共{len(oracle)}条；边界测试通过。", flush=True)
    base.dump_json(run / "correctness.json", checks)
    base.write_csv(run / "bbox_candidate_counts.csv", counts)
    special = next(r for r in counts if r["dataset"] == "boundary_A" and r["resolution"] == 7)
    assert (special["border_candidates"], special["bbox_removed"], special["new_exact_points"]) == (8198, 2965, 5233)
    base.dump_json(run / "boundary_A_r7_audit.json", special)
    return checks, oracles, counts


def measure_batch(method, data, k):
    query_function = query_h3_bbox if method.endswith("_bbox") else base.query
    start = time.perf_counter_ns()
    state, preparation = base.prepare_method(original_method(method), data)
    after_preparation = time.perf_counter_ns()
    for _ in range(k):
        ids, metrics = query_function(state, data)
    end = time.perf_counter_ns()
    return ids, {"whole_total_ns": end - start, "preparation_external_ns": after_preparation - start,
                 "continuous_queries_ns": end - after_preparation, "query_calls": k, **preparation,
                 **{f"last_{key}": value for key, value in metrics.items()}}


def benchmark(run, datasets, oracles, protocol):
    records = []
    cases = [(name, method) for name in protocol["datasets"] for method in protocol["methods"]]
    randomizer = random.Random(protocol["order_seed"])
    gc_enabled = gc.isenabled()
    gc.disable()
    try:
        for name, method in cases:
            state, _ = base.prepare_method(original_method(method), datasets[name])
            run_query(method, state, datasets[name])
        del state
        for trial in range(1, protocol["whole_batch_repeats"] + 1):
            order = cases.copy()
            randomizer.shuffle(order)
            for position, (name, method) in enumerate(order, 1):
                ids, metrics = measure_batch(method, datasets[name], protocol["K"])
                actual = set(map(int, ids))
                if actual != oracles[name] or len(ids) != len(actual):
                    base.dump_json(run / "batch_failure.json", {"dataset": name, "method": method,
                                   "trial": trial, "extra": sorted(actual - oracles[name]),
                                   "missing": sorted(oracles[name] - actual), "metrics": metrics})
                    if records:
                        base.write_csv(run / "whole_batch_trials_before_failure.csv", records)
                    raise RuntimeError("计时后正确性检查失败")
                records.append({"trial": trial, "order_in_round": position, "dataset": name,
                                "method": method, "K": protocol["K"], **metrics})
            print(f"整体交错计时：{trial}/{protocol['whole_batch_repeats']}轮完成", flush=True)
    finally:
        if gc_enabled:
            gc.enable()
    return records


def profiling(datasets, oracles, protocol):
    records = []
    cases = [(name, method) for name in protocol["datasets"] for method in protocol["methods"]]
    states = {(name, method): base.prepare_method(original_method(method), datasets[name])[0]
              for name, method in cases}
    randomizer = random.Random(protocol["order_seed"] + 1)
    gc_enabled = gc.isenabled()
    gc.disable()
    try:
        for name, method in cases:
            for _ in range(3):
                run_query(method, states[name, method], datasets[name], profile=True)
        for trial in range(1, protocol["profile_repeats"] + 1):
            order = cases.copy()
            randomizer.shuffle(order)
            for position, (name, method) in enumerate(order, 1):
                ids, metrics = run_query(method, states[name, method], datasets[name], profile=True)
                assert set(map(int, ids)) == oracles[name] and len(ids) == len(oracles[name])
                # 原版没有新增矩形步骤；这里以0记录，原始函数和计时点保持不变。
                for key in ("bbox_mask_ns", "bbox_array_compaction_ns", "bbox_filter_ns"):
                    metrics.setdefault(key, 0)
                metrics.setdefault("pre_bbox_gather_ns", metrics["lookup_and_gather_ns"])
                metrics.setdefault("border_candidates_before_bbox", metrics["exact_points"])
                metrics.setdefault("bbox_removed_points", 0)
                records.append({"trial": trial, "order_in_round": position, "dataset": name,
                                "method": method, **metrics})
    finally:
        if gc_enabled:
            gc.enable()
    return records


def summarize(whole, profiles, counts, protocol):
    totals, phase_summary, pairs, rankings = [], [], [], []
    for name in protocol["datasets"]:
        for method in protocol["methods"]:
            rows = [r for r in whole if r["dataset"] == name and r["method"] == method]
            totals.append({"dataset": name, "method": method, "trials": len(rows),
                           **distribution_stats([r["whole_total_ns"] for r in rows]),
                           "prepare_median_ms": statistics.median(r["preparation_external_ns"] for r in rows) / 1e6,
                           "queries_median_ms": statistics.median(r["continuous_queries_ns"] for r in rows) / 1e6})
            phase_rows = [r for r in profiles if r["dataset"] == name and r["method"] == method]
            metrics = [k for k in phase_rows[0] if k.endswith("_ns")]
            phase_summary.append({"dataset": name, "method": method, "trials": len(phase_rows),
                                  **{k.removesuffix("_ns") + "_median_us": statistics.median(r[k] for r in phase_rows) / 1000
                                     for k in metrics},
                                  **distribution_stats([r["query_total_ns"] for r in phase_rows])})
        for resolution in (6, 7, 8):
            original = f"h3_r{resolution}"
            filtered = original + "_bbox"
            old = next(r for r in totals if r["dataset"] == name and r["method"] == original)
            new = next(r for r in totals if r["dataset"] == name and r["method"] == filtered)
            oldp = next(r for r in phase_summary if r["dataset"] == name and r["method"] == original)
            newp = next(r for r in phase_summary if r["dataset"] == name and r["method"] == filtered)
            count = next(r for r in counts if r["dataset"] == name and r["resolution"] == resolution)
            delta = old["median_ms"] - new["median_ms"]
            pairs.append({**count, "bbox_filter_median_us": newp["bbox_filter_median_us"],
                          "bbox_mask_median_us": newp["bbox_mask_median_us"],
                          "bbox_compaction_median_us": newp["bbox_array_compaction_median_us"],
                          "original_exact_median_us": oldp["exact_predicate_median_us"],
                          "new_exact_median_us": newp["exact_predicate_median_us"],
                          "exact_saved_us": oldp["exact_predicate_median_us"] - newp["exact_predicate_median_us"],
                          "original_profile_query_us": oldp["query_total_median_us"],
                          "new_profile_query_us": newp["query_total_median_us"],
                          "original_total_ms": old["median_ms"], "new_total_ms": new["median_ms"],
                          "total_saved_ms": delta, "total_saved_percent": 100 * delta / old["median_ms"],
                          "original_q25_ms": old["q25_ms"], "original_q75_ms": old["q75_ms"],
                          "new_q25_ms": new["q25_ms"], "new_q75_ms": new["q75_ms"],
                          "iqr_overlap": not (new["q75_ms"] < old["q25_ms"] or old["q75_ms"] < new["q25_ms"])})
        original_order = sorted([r for r in totals if r["dataset"] == name and not r["method"].endswith("_bbox")],
                                key=lambda r: r["median_ms"])
        new_order = sorted([r for r in totals if r["dataset"] == name and (r["method"] == "direct" or r["method"].endswith("_bbox"))],
                           key=lambda r: r["median_ms"])
        all_order = sorted([r for r in totals if r["dataset"] == name], key=lambda r: r["median_ms"])
        rankings.append({"dataset": name, "original_ranking": " < ".join(r["method"] for r in original_order),
                         "bbox_ranking": " < ".join(r["method"] for r in new_order),
                         "all_seven_ranking": " < ".join(r["method"] for r in all_order),
                         "original_best": original_order[0]["method"], "original_best_ms": original_order[0]["median_ms"],
                         "bbox_best": new_order[0]["method"], "bbox_best_ms": new_order[0]["median_ms"]})
    return totals, phase_summary, pairs, rankings


def write_results(run, totals, pairs, rankings, checks):
    lines = ["# H3边界候选增加外包矩形筛选：单因素对照", "",
             "使用上一轮四组保存的原始模拟数据，未重新生成或调整分布。原查询函数不修改，新增版只对边界候选加含边界矩形筛选。", "",
             "核实：boundary_A/r7共有8198个边界候选，其中2965个在海区外包矩形外，筛后剩5233个精判点。", "",
             "## 候选数量、筛选费用与总时间", "",
             "整体时间是一次新准备加连续1000次查询的实测中位数，每种7次、同轮随机交错。分项来自单独的31轮交错诊断，单位微秒/单次查询，不将分项中位数相加替代总计。", "",
             "| 数据/档位 | 原边界候选 | 矩形排除 | 新精判点 | 筛选费μs | 精判节省μs | 原总计ms | 新总计ms | 总节省% |",
             "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |"]
    for r in pairs:
        lines.append(f"| {r['dataset']}/r{r['resolution']} | {r['border_candidates']} | {r['bbox_removed']} | "
                     f"{r['new_exact_points']} | {r['bbox_filter_median_us']:.2f} | {r['exact_saved_us']:.2f} | "
                     f"{r['original_total_ms']:.3f} | {r['new_total_ms']:.3f} | {r['total_saved_percent']:.2f} |")
    lines += ["", "总节省为正表示变快，为负表示变慢。筛选费包括矩形条件计算及行号、经度、纬度三个数组的同步紧凑化；初始坐标读取也始终计入查询总时间。",
              "直接查询的新增bbox费用记0，表示没有新增H3边界筛选步骤；直接方法原本的矩形筛选费用仍在lookup_and_gather内。",
              "独立诊断中新版额外启用过滤计时探针，主实验没有这些额外探针。由于缓存状态、准备耗时和探针存在差异，不用诊断分项乘1000解释精确的整体差额。", "",
              "## 排名（由快到慢）", "",
              "| 数据 | 原四方案排名 | 加矩形后四方案排名 | 同轮直接查询总计ms |",
              "| --- | --- | --- | ---: |"]
    for r in rankings:
        direct = next(s for s in totals if s["dataset"] == r["dataset"] and s["method"] == "direct")
        lines.append(f"| {r['dataset']} | {r['original_ranking']} | {r['bbox_ranking']} | {direct['median_ms']:.3f} |")
    lines += ["", "## 正确性", "",
              "四组的原版及新增版均与直接精判参照逐条核对记录ID；所有7方案误选、漏选、重复均为0。每组32个边界测试点通过，内部格直接纳入逻辑不变。",
              "原版保守覆盖、相交格分类及全部正确性审计沿用旧实现。每个整体试次后在计时之外核对最后结果，计时中输入只读且每次重新构造返回数组。", "",
              "## 总时间的波动", "",
              "| 数据 | 方法 | 中位数ms | 第25–75百分位ms | 最小–最大ms |",
              "| --- | --- | ---: | --- | --- |"]
    for r in totals:
        lines.append(f"| {r['dataset']} | {r['method']} | {r['median_ms']:.4f} | "
                     f"{r['q25_ms']:.4f}–{r['q75_ms']:.4f} | {r['min_ms']:.4f}–{r['max_ms']:.4f} |")
    lines += ["", "文件：whole_batch_trials.csv（196条原始整体记录）、profile_trials.csv（868条诊断记录）、pair_comparisons.csv、rankings.csv、correctness.json、bbox_candidate_counts.csv。",
              "protocol.json记录计时协议；input_data/保存逐字节复制的旧输入，input_manifest.json记录来源与指纹；source_snapshots/保留新旧代码。",
              "这轮只检验该实现因素，不删除旧结果，不实现选择规则。", ""]
    (run / "矩形筛选对照结果.md").write_text("\n".join(lines), encoding="utf-8")


def existing_files():
    files = [p for p in ROOT.iterdir() if p.is_file()]
    files += [p for p in (ROOT / "runs").rglob("*") if p.is_file()]
    return {p.relative_to(ROOT).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}


def main():
    protocol_bytes = PROTOCOL_PATH.read_bytes()
    protocol = json.loads(protocol_bytes)
    protected = existing_files()
    source = ROOT / protocol["source_run"]
    assert protocol["K"] == 1000 and protocol["whole_batch_repeats"] >= 7
    run = ROOT / "runs" / datetime.now().astimezone().strftime("bbox_%Y%m%d_%H%M%S_%f")
    run.mkdir(parents=True)
    (run / "protocol.json").write_bytes(protocol_bytes)
    base.dump_json(run / "protocol_record.json", {"before_timing": datetime.now().astimezone().isoformat(),
                                                 "sha256": hashlib.sha256(protocol_bytes).hexdigest()})
    base.dump_json(run / "previous_files_sha256.json", protected)
    for name in ("bbox_ablation.py", "test_bbox_ablation.py", "run_experiment.py", "requirements.txt"):
        path = run / "source_snapshots" / name
        path.parent.mkdir(exist_ok=True)
        path.write_bytes((ROOT / name).read_bytes())
    environment = base.environment_info()
    environment["method_order_seed"] = protocol["order_seed"]
    base.dump_json(run / "environment.json", environment)
    datasets, input_manifest = {}, {}
    for name in protocol["datasets"]:
        folder = source / "datasets" / name
        datasets[name] = load_saved_data(folder)
        for filename in ("simulated_points.csv", "simulated_sea.geojson"):
            src = folder / "data" / filename
            dst = run / "input_data" / name / filename
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            digest = hashlib.sha256(src.read_bytes()).hexdigest()
            assert digest == hashlib.sha256(dst.read_bytes()).hexdigest()
            input_manifest[dst.relative_to(run).as_posix()] = {"source": src.relative_to(ROOT).as_posix(), "sha256": digest}
    base.dump_json(run / "input_manifest.json", input_manifest)
    print(f"矩形筛选对照目录：{run}", flush=True)
    checks, oracles, counts = preflight(run, source, datasets, protocol)
    whole = benchmark(run, datasets, oracles, protocol)
    base.write_csv(run / "whole_batch_trials.csv", whole)
    print("开始单独的筛选/精判分项诊断。", flush=True)
    profiles = profiling(datasets, oracles, protocol)
    base.write_csv(run / "profile_trials.csv", profiles)
    totals, phases, pairs, rankings = summarize(whole, profiles, counts, protocol)
    for filename, records in (("whole_batch_summary.csv", totals), ("profile_summary.csv", phases),
                               ("pair_comparisons.csv", pairs), ("rankings.csv", rankings)):
        base.write_csv(run / filename, records)
    base.dump_json(run / "summary.json", {"protocol": protocol, "correctness": checks,
                                         "totals": totals, "phases": phases, "pairs": pairs, "rankings": rankings})
    write_results(run, totals, pairs, rankings, checks)
    assert PROTOCOL_PATH.read_bytes() == protocol_bytes
    assert all(hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == digest for name, digest in protected.items())
    base.dump_json(run / "status.json", {"status": "completed", "whole_trials": len(whole),
                                        "profile_trials": len(profiles), "queries_in_main": sum(r["K"] for r in whole),
                                        "old_files_unchanged": True, "data_regenerated": False})
    base.dump_json(run / "sha256_manifest.json", {p.relative_to(run).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
                                                  for p in run.rglob("*") if p.is_file()})
    (ROOT / "LATEST_BBOX_ABLATION.txt").write_text(run.relative_to(ROOT).as_posix() + "\n", encoding="utf-8")
    for r in pairs:
        print(f"{r['dataset']}/r{r['resolution']}: {r['border_candidates']}→{r['new_exact_points']}点；"
              f"总计{r['original_total_ms']:.3f}→{r['new_total_ms']:.3f}ms，节省{r['total_saved_percent']:.2f}%", flush=True)


if __name__ == "__main__":
    main()
