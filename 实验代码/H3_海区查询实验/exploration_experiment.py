"""四组探索数据：实测一次准备+连续K次查询，以及200点样本的部分特征成本。

保留并导入原 run_experiment.py 的查询/准备实现，不修改上一轮运行。
运行：python exploration_experiment.py
"""
from __future__ import annotations

import argparse
from datetime import datetime
import gc
import hashlib
import json
from pathlib import Path
import random
import statistics
import time

import run_experiment as base

np, shapely, h3 = base.np, base.shapely, base.h3
ROOT = Path(__file__).resolve().parent
PROTOCOL_PATH = ROOT / "exploration_protocol.json"


def generate_dataset(case: dict, protocol: dict):
    n = protocol["random_points_per_dataset"]
    if case["distribution"] == "uniform":
        return base.generate_data(n, case["seed"]), {"uniform_component_points": n}
    cfg = protocol["boundary_cluster"]
    rng = np.random.Generator(np.random.PCG64(case["seed"]))
    region = shapely.Polygon(base.VERTICES)
    boundary = shapely.boundary(region)
    xmin, ymin, xmax, ymax = base.EXTENT
    uniform_n = cfg["uniform_component_points"]
    uniform = np.column_stack((rng.uniform(xmin, xmax, uniform_n), rng.uniform(ymin, ymax, uniform_n)))
    inside_chunks, outside_chunks = [], []
    inside_n = outside_n = proposals = 0
    for _ in range(1000):
        batch = cfg["proposal_batch_size"]
        xy = np.column_stack((rng.uniform(xmin, xmax, batch), rng.uniform(ymin, ymax, batch)))
        near = shapely.distance(boundary, shapely.points(xy)) <= cfg["max_distance_to_region_boundary_degrees"]
        band = xy[near]
        covered = shapely.intersects_xy(region, band[:, 0], band[:, 1])
        take_inside = band[covered][:max(0, cfg["inside_band_points"] - inside_n)]
        take_outside = band[~covered][:max(0, cfg["outside_band_points"] - outside_n)]
        inside_chunks.append(take_inside)
        outside_chunks.append(take_outside)
        inside_n += len(take_inside)
        outside_n += len(take_outside)
        proposals += batch
        if inside_n == cfg["inside_band_points"] and outside_n == cfg["outside_band_points"]:
            break
    else:
        raise RuntimeError("边界拒绝采样未满足事先规定的数量；停止，不调整生成参数")
    inside, outside = np.concatenate(inside_chunks), np.concatenate(outside_chunks)
    assert np.all(shapely.intersects_xy(region, inside[:, 0], inside[:, 1]))
    assert not np.any(shapely.intersects_xy(region, outside[:, 0], outside[:, 1]))
    xy = np.concatenate((uniform, inside, outside))
    categories = (["uniform_component"] * uniform_n + ["boundary_band_inside"] * inside_n
                  + ["boundary_band_outside"] * outside_n)
    assert len(xy) == n
    order = rng.permutation(n)
    xy = xy[order]
    categories = [categories[int(i)] for i in order]
    probes = base.generate_data(0, case["seed"])
    data = base.Dataset(np.arange(1, n + len(probes.ids) + 1, dtype=np.int64),
                        np.r_[xy[:, 0], probes.lon], np.r_[xy[:, 1], probes.lat],
                        categories + probes.category,
                        [None] * n + probes.expected_probe_membership)
    stats = {"uniform_component_points": uniform_n, "inside_band_component_points": inside_n,
             "outside_band_component_points": outside_n, "rejection_proposals": proposals}
    return data, stats


def save_dataset(folder: Path, data, case, protocol, generation):
    rows = [{"record_id": int(record_id), "longitude": format(lon, ".17g"),
             "latitude": format(lat, ".17g"), "category": label,
             "probe_expected_inside": "" if truth is None else int(truth)}
            for record_id, lon, lat, label, truth in zip(data.ids, data.lon, data.lat,
                                                        data.category, data.expected_probe_membership)]
    base.write_csv(folder / "data" / "simulated_points.csv", rows)
    base.dump_json(folder / "data" / "simulated_sea.geojson", {
        "type": "Feature", "properties": {"simulated": True},
        "geometry": shapely.Polygon(base.VERTICES).__geo_interface__})
    base.dump_json(folder / "data" / "generation_rules.json", {
        "case": case, "generation": generation, "all_data_simulated": True,
        "protocol": protocol, "geometry_model": "same local lon/lat planar straight segments as round 1",
        "boundary_probes_identical_to_previous_round": True})


def feature_extraction(data, prepared_region, sample_size: int, sample_seed: int):
    """只使用原始点和共用几何，不接受全量点索引或海区覆盖清单。

    返回样本行号和特征计时；不实现成本预测或方法选择。
    """
    start = time.perf_counter_ns()
    rng = np.random.Generator(np.random.PCG64(sample_seed))
    rows = rng.choice(len(data.ids), sample_size, replace=False)
    x, y = data.lon[rows], data.lat[rows]
    after_sampling = time.perf_counter_ns()
    metrics = {"sample_and_gather_ns": after_sampling - start}
    for resolution in base.RESOLUTIONS:
        rstart = time.perf_counter_ns()
        cells = np.fromiter((h3.latlng_to_cell(float(lat), float(lon), resolution)
                             for lon, lat in zip(x, y)), dtype=np.uint64, count=sample_size)
        after_encoding = time.perf_counter_ns()
        unique = tuple(map(int, np.unique(cells)))
        after_deduplication = time.perf_counter_ns()
        cover, internal, border, _ = base.classify_cells(prepared_region, unique)
        after_classification = time.perf_counter_ns()
        metrics.update({f"r{resolution}_encoding_ns": after_encoding - rstart,
                        f"r{resolution}_deduplication_ns": after_deduplication - after_encoding,
                        f"r{resolution}_polygon_and_classification_ns": after_classification - after_deduplication,
                        f"r{resolution}_unique_cells": len(unique),
                        f"r{resolution}_sample_cells_internal": len(internal),
                        f"r{resolution}_sample_cells_border": len(border),
                        f"r{resolution}_sample_cells_outside": len(unique) - len(cover)})
    end = time.perf_counter_ns()
    metrics["feature_only_ns"] = end - start
    metrics["feature_other_ns"] = end - start - sum(v for k, v in metrics.items()
                                                  if k.endswith("_ns") and k != "feature_only_ns")
    return rows, metrics


def measure_batch(method, data, k):
    """唯一外层计时：准备+连续查询。循环内无校验、文件写入或最终答案缓存。"""
    start = time.perf_counter_ns()
    state, preparation = base.prepare_method(method, data)
    after_preparation = time.perf_counter_ns()
    for _ in range(k):
        ids, last_query_metrics = base.query(state, data)
    end = time.perf_counter_ns()
    metrics = {"whole_total_ns": end - start, "preparation_external_ns": after_preparation - start,
               "continuous_queries_ns": end - after_preparation, "query_calls": k, **preparation,
               **{f"last_{key}": value for key, value in last_query_metrics.items()}}
    # ID 集合比较由调用方在此函数返回之后进行，明确在计时之外。
    return ids, metrics


def benchmark_all(datasets, oracles, protocol, run):
    records = []
    cases = [(case["name"], k, method) for case in protocol["datasets"]
             for k in protocol["K"] for method in protocol["methods"]]
    randomizer = random.Random(protocol["method_case_order_seed"])
    gc_enabled = gc.isenabled()
    gc.disable()
    try:
        for name, data in datasets.items():
            for method in protocol["methods"]:
                state, _ = base.prepare_method(method, data)
                base.query(state, data)
        del state
        for trial in range(1, protocol["whole_batch_repeats"] + 1):
            order = cases.copy()
            randomizer.shuffle(order)
            for position, (name, k, method) in enumerate(order, 1):
                ids, metrics = measure_batch(method, datasets[name], k)
                actual = set(map(int, ids))
                if actual != oracles[name] or len(ids) != len(actual):
                    diagnostics = {"dataset": name, "method": method, "K": k, "trial": trial,
                                   "false_positive_ids": sorted(actual - oracles[name]),
                                   "false_negative_ids": sorted(oracles[name] - actual),
                                   "duplicate_ids": len(ids) - len(actual), "failed_batch_metrics": metrics}
                    base.dump_json(run / "post_batch_failure.json", diagnostics)
                    if records:
                        base.write_csv(run / "whole_batch_trials_before_failure.csv", records)
                    raise RuntimeError(f"计时后的 ID 检查失败：{name}/{method}/K={k}")
                records.append({"trial": trial, "order_in_round": position, "dataset": name,
                                "K": k, "method": method, **metrics})
            print(f"连续执行计时：第 {trial}/{protocol['whole_batch_repeats']} 轮完成", flush=True)
    finally:
        if gc_enabled:
            gc.enable()
    return records


def benchmark_features(datasets, protocol):
    records = []
    randomizer = random.Random(protocol["method_case_order_seed"] + 1)
    cases = {case["name"]: case for case in protocol["datasets"]}
    gc_enabled = gc.isenabled()
    gc.disable()
    try:
        for name, data in datasets.items():
            feature_extraction(data, base.make_region(), protocol["feature_sample_size"],
                               protocol["feature_seed_base"] + cases[name]["seed"])
        for trial in range(1, protocol["feature_repeats"] + 1):
            names = list(datasets)
            randomizer.shuffle(names)
            for position, name in enumerate(names, 1):
                sample_seed = protocol["feature_seed_base"] + cases[name]["seed"] * 1000 + trial
                common_start = time.perf_counter_ns()
                region = base.make_region()
                common_end = time.perf_counter_ns()
                rows, metrics = feature_extraction(datasets[name], region, protocol["feature_sample_size"], sample_seed)
                records.append({"trial": trial, "order_in_round": position, "dataset": name,
                                "sample_seed": sample_seed, "sample_size": protocol["feature_sample_size"],
                                "common_geometry_prepare_ns": common_end - common_start, **metrics,
                                "sample_record_ids": json.dumps(datasets[name].ids[rows].tolist())})
    finally:
        if gc_enabled:
            gc.enable()
    return records


def distribution_stats(values):
    values = np.asarray(values, dtype=float) / 1e6
    return {"median_ms": float(np.median(values)), "q25_ms": float(np.percentile(values, 25)),
            "q75_ms": float(np.percentile(values, 75)), "min_ms": float(np.min(values)),
            "max_ms": float(np.max(values)), "mad_ms": float(np.median(np.abs(values - np.median(values))))}


def summarize(records, feature_records, protocol):
    summaries, feature_summaries, comparisons = [], [], []
    for case in protocol["datasets"]:
        name = case["name"]
        feature = [row for row in feature_records if row["dataset"] == name]
        item = {"dataset": name, "distribution": case["distribution"], "seed": case["seed"],
                "trials": len(feature), **distribution_stats([row["feature_only_ns"] for row in feature])}
        ns_columns = [key for key in feature[0] if key.endswith("_ns") and key != "feature_only_ns"]
        item.update({key.removesuffix("_ns") + "_median_ms": statistics.median(row[key] for row in feature) / 1e6
                     for key in ns_columns})
        feature_summaries.append(item)
        for k in protocol["K"]:
            options = []
            for method in protocol["methods"]:
                raw = [row for row in records if row["dataset"] == name and row["K"] == k and row["method"] == method]
                item = {"dataset": name, "distribution": case["distribution"], "seed": case["seed"],
                        "K": k, "method": method, "trials": len(raw),
                        **distribution_stats([row["whole_total_ns"] for row in raw]),
                        "prepare_median_ms": statistics.median(row["preparation_external_ns"] for row in raw) / 1e6,
                        "continuous_queries_median_ms": statistics.median(row["continuous_queries_ns"] for row in raw) / 1e6}
                summaries.append(item)
                options.append(item)
            two = min((row for row in options if row["method"] in ("direct", "h3_r7")), key=lambda row: row["median_ms"])
            four = min(options, key=lambda row: row["median_ms"])
            saving = two["median_ms"] - four["median_ms"]
            feature_cost = feature_summaries[-1]["median_ms"]
            comparisons.append({"dataset": name, "distribution": case["distribution"], "seed": case["seed"],
                                "K": k, "two_choice_method": two["method"], "two_choice_ms": two["median_ms"],
                                "two_choice_q25_ms": two["q25_ms"], "two_choice_q75_ms": two["q75_ms"],
                                "two_choice_min_ms": two["min_ms"], "two_choice_max_ms": two["max_ms"],
                                "four_choice_method": four["method"], "four_choice_ms": four["median_ms"],
                                "four_choice_q25_ms": four["q25_ms"], "four_choice_q75_ms": four["q75_ms"],
                                "four_choice_min_ms": four["min_ms"], "four_choice_max_ms": four["max_ms"],
                                "potential_saving_ms": saving, "potential_saving_percent": 100 * saving / two["median_ms"],
                                "feature_only_median_ms": feature_cost, "saving_over_feature_cost": saving / feature_cost,
                                "winning_iqr_disjoint": four["q75_ms"] < two["q25_ms"]})
    return summaries, feature_summaries, comparisons


def write_results(run, comparisons, summaries, feature_summaries, checks, protocol):
    label = {"direct": "直接", "h3_r6": "r6", "h3_r7": "r7", "h3_r8": "r8"}
    lines = ["# 四组探索数据：准备与连续重算对照", "",
             "海区、船位均为模拟数据；本轮全部是探索数据，不是选择规则的验证集。没有实现选择器。", "",
             "固定同一海区，每组 10,000 个随机点和原 32 个边界测试点。K 表示实验要求重算的次数；最终答案不缓存。", "",
             "以下是外层计时器直接测得的一次准备加连续 K 次查询的总耗时，不是单次耗时乘 K 的估算。",
             f"每种情况 {protocol['whole_batch_repeats']} 次；单位毫秒。括号为第25–75百分位区间（中间一半记录所在的范围）。", "",
             "| 数据 | K | 二选一事后最优：中位数（波动区间） | 四选一事后最优：中位数（波动区间） | 潜在节省 | 200点样本费用 | 节省/样本费 |",
             "| --- | ---: | --- | --- | ---: | ---: | ---: |"]
    for r in comparisons:
        lines.append(f"| {r['dataset']} | {r['K']} | {label[r['two_choice_method']]} "
                     f"{r['two_choice_ms']:.3f}（{r['two_choice_q25_ms']:.3f}–{r['two_choice_q75_ms']:.3f}） | "
                     f"{label[r['four_choice_method']]} {r['four_choice_ms']:.3f}"
                     f"（{r['four_choice_q25_ms']:.3f}–{r['four_choice_q75_ms']:.3f}） | "
                     f"{r['potential_saving_ms']:.3f} | {r['feature_only_median_ms']:.3f} | "
                     f"{r['saving_over_feature_cost']:.2f} |")
    lines += ["", "二选一的候选是直接查询、r7；四选一增加 r6、r8。先求各方法总耗时的中位数，再事后选最小值。",
              "潜在节省是两个事后最优中位数之差，属于事后可选方法范围扩大后的潜在空间，不是已实现选择器的收益。", "",
              "## 正确性", "",
              "| 数据 | 种子 | 四方案共同返回数 | 各方案误选/漏选 | 32个边界点 | 全覆盖审计 |",
              "| --- | ---: | ---: | --- | --- | --- |"]
    for case in protocol["datasets"]:
        c = checks[case["name"]]
        lines.append(f"| {case['name']} | {case['seed']} | {c['oracle_selected_points']} | "
                     f"全部0/0 | {'通过' if c['probe_checks_passed'] else '失败'} | "
                     f"{'通过' if c['all_cover_checks_passed'] else '失败'} |")
    lines += ["", "## 样本部分成本", "",
              "只计随机抽200条、三档分别编码、各档样本去重格子构造轮廓并分类。没有借用全量索引或完整覆盖清单。",
              "每次独立采样，不放回；每次种子与200个记录ID保存到 feature_trials.csv。样本计时每组31次。", "",
              "| 数据 | 特征提取中位数 | 中间一半区间 | 最小–最大 | 共用几何准备中位数 |",
              "| --- | ---: | --- | --- | ---: |"]
    for r in feature_summaries:
        lines.append(f"| {r['dataset']} | {r['median_ms']:.4f} | {r['q25_ms']:.4f}–{r['q75_ms']:.4f} | "
                     f"{r['min_ms']:.4f}–{r['max_ms']:.4f} | {r['common_geometry_prepare_median_ms']:.4f} |")
    lines += ["", "共用几何准备与样本特征分别计时，没有混入特征列；实际选择流程若已完成并保有同一预处理几何，可复用，不重复计入。",
              "样本列只包含选择过程的一部分。尚未包含完整成本预测、决策、样本结果复用和最终方法准备等完整记账。", "",
              "## 四种方法的完整波动", "",
              "| 数据 | K | 方法 | 总耗时中位数 | 第25–75百分位 | 最小–最大 |",
              "| --- | ---: | --- | ---: | --- | --- |"]
    for r in summaries:
        lines.append(f"| {r['dataset']} | {r['K']} | {label[r['method']]} | {r['median_ms']:.4f} | "
                     f"{r['q25_ms']:.4f}–{r['q75_ms']:.4f} | {r['min_ms']:.4f}–{r['max_ms']:.4f} |")
    lines += ["", "## 口径与复现", "",
              "- 新运行目录独立保存，不覆盖上轮记录；代码仍调用原来的 prepare_method 和 query。",
              "- 每个整体试次都重新建立索引、生成清单并分类；没有跨试次复用这些准备资产。函数与库预热一致。",
              "- 外层整体计时包含原查询函数自身的计时与指标字典开销、循环和每次新建返回数组的开销。",
              "- 预检完整ID集合、边界点及覆盖审计；输入数组在计时中只读，每个整体试次结束后再核对最后一次返回ID。",
              "- 循环内不做正确性集合比较，不保留全部查询结果，也不以已有答案代替查询。",
              "- 数据生成、文件写入、正确性校验、覆盖审计、依赖导入均在计时之外。所有案例每轮随机打乱，GC统一暂停。",
              "- 边界聚集：80%为距边界不超过0.002度的点，内外各4000；20%为背景均匀点；最后加32个测试点。",
              "- 0.002度是本实验经纬度平面距离，不将其解释为精确固定米数。背景点也可能恰好落入边界带。",
              "- 四组选用新种子；协议在生成数据和性能计时之前写盘并记录SHA-256，未根据结果改参数。",
              "- 事后最优利用了完整计时信息，只作探索参照，后续选择器不得提前读取验证数据上的这个答案。", "",
              "文件：whole_batch_trials.csv、whole_batch_summary.csv、posthoc_comparisons.csv、feature_trials.csv、feature_summary.csv。",
              "协议见 protocol.json；每组原始模拟点、覆盖、编码、正确性及最终ID集合在 datasets/ 下。", ""]
    (run / "探索结果.md").write_text("\n".join(lines), encoding="utf-8")
    (ROOT / "探索对照结果.md").write_text("\n".join(lines[:8]) + "\n\n完整结果与波动表："
        f"[本次探索结果]({run.relative_to(ROOT).as_posix()}/探索结果.md)。\n", encoding="utf-8")


def protected_files():
    fixed = ("run_experiment.py", "test_experiment.py", "requirements.txt", "README.md", "run.ps1",
             "LATEST_RUN.txt", "结果速览.md", "本轮解读.md")
    files = [ROOT / name for name in fixed if (ROOT / name).is_file()]
    files += [p for p in (ROOT / "runs").rglob("*") if p.is_file()]
    return {p.relative_to(ROOT).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.parse_args()
    protocol = json.loads(PROTOCOL_PATH.read_text(encoding="utf-8"))
    assert protocol["methods"] == list(base.METHODS)
    assert protocol["region_vertices_lon_lat"] == [list(v) for v in base.VERTICES]
    assert protocol["extent_lon_lat"] == list(base.EXTENT)
    assert protocol["whole_batch_repeats"] >= 5 and protocol["feature_repeats"] >= 5
    assert len({case["seed"] for case in protocol["datasets"]}) == 4
    protected = protected_files()
    run = ROOT / "runs" / datetime.now().astimezone().strftime("explore_%Y%m%d_%H%M%S_%f")
    run.mkdir(parents=True)
    protocol_bytes = PROTOCOL_PATH.read_bytes()
    (run / "protocol.json").write_bytes(protocol_bytes)
    base.dump_json(run / "protocol_record.json", {
        "recorded_before_generation_and_timing": datetime.now().astimezone().isoformat(),
        "source_protocol_creation_time_ns": PROTOCOL_PATH.stat().st_ctime_ns,
        "protocol_sha256": hashlib.sha256(protocol_bytes).hexdigest()})
    base.dump_json(run / "previous_files_sha256.json", protected)
    for name in ("exploration_experiment.py", "run_experiment.py", "requirements.txt", "test_exploration.py"):
        (run / name).write_bytes((ROOT / name).read_bytes())
    environment = base.environment_info()
    environment["method_order_seed"] = protocol["method_case_order_seed"]
    environment["feature_seed_base"] = protocol["feature_seed_base"]
    base.dump_json(run / "environment.json", environment)
    print(f"新探索运行：{run}", flush=True)
    print(f"协议 SHA-256：{hashlib.sha256(protocol_bytes).hexdigest()}", flush=True)
    datasets, oracles, checks = {}, {}, {}
    for case in protocol["datasets"]:
        name = case["name"]
        data, generation = generate_dataset(case, protocol)
        assert len(data.ids) == 10032
        save_dataset(run / "datasets" / name, data, case, protocol, generation)
        check, oracle = base.correctness_phase(run / "datasets" / name, data)
        for array in (data.ids, data.lon, data.lat):
            array.flags.writeable = False
        datasets[name], oracles[name], checks[name] = data, oracle, check
        print(f"{name}：正确性及32个边界点通过；4种方法共同返回 {len(oracle)} 条。", flush=True)
    base.dump_json(run / "all_correctness.json", checks)
    whole_records = benchmark_all(datasets, oracles, protocol, run)
    base.write_csv(run / "whole_batch_trials.csv", whole_records)
    print("开始独立测量200条样本特征；不借用全量索引或覆盖清单。", flush=True)
    feature_records = benchmark_features(datasets, protocol)
    base.write_csv(run / "feature_trials.csv", feature_records)
    summaries, feature_summaries, comparisons = summarize(whole_records, feature_records, protocol)
    base.write_csv(run / "whole_batch_summary.csv", summaries)
    base.write_csv(run / "feature_summary.csv", feature_summaries)
    base.write_csv(run / "posthoc_comparisons.csv", comparisons)
    base.dump_json(run / "summary.json", {"exploratory_only": True, "all_data_simulated": True,
                                         "whole_batch_results_are_measured_not_extrapolated": True,
                                         "protocol": protocol, "correctness": checks,
                                         "method_summary": summaries, "feature_summary": feature_summaries,
                                         "posthoc_comparisons": comparisons})
    write_results(run, comparisons, summaries, feature_summaries, checks, protocol)
    assert PROTOCOL_PATH.read_bytes() == protocol_bytes, "运行后协议发生变化，停止发布本次摘要"
    assert all(hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == digest for name, digest in protected.items())
    base.dump_json(run / "status.json", {"status": "completed", "previous_files_unchanged": True,
                                        "whole_batch_trials": len(whole_records), "feature_trials": len(feature_records),
                                        "queries_recomputed_in_measured_batches": sum(row["K"] for row in whole_records)})
    manifest = {p.relative_to(run).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
                for p in run.rglob("*") if p.is_file()}
    base.dump_json(run / "sha256_manifest.json", manifest)
    (ROOT / "LATEST_EXPLORATION.txt").write_text(run.relative_to(ROOT).as_posix() + "\n", encoding="utf-8")
    for r in comparisons:
        print(f"{r['dataset']}, K={r['K']}: 二选一 {r['two_choice_method']} {r['two_choice_ms']:.3f} ms；"
              f"四选一 {r['four_choice_method']} {r['four_choice_ms']:.3f} ms；"
              f"潜在节省 {r['potential_saving_ms']:.3f} ms；样本特征 {r['feature_only_median_ms']:.3f} ms")
    print("探索测量完成，上一轮记录保持不变。", flush=True)


if __name__ == "__main__":
    main()
