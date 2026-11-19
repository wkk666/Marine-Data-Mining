"""旧四组统一入口校准；不创建或读取六组独立验证数据。"""
from datetime import datetime
import gc
import hashlib
import json
from pathlib import Path
import random
import shutil
import statistics
import time

import bbox_ablation as previous
import selector_engine as engine
from selector_decision import decision_intervals
from exploration_experiment import distribution_stats

base, np = engine.base, engine.np
ROOT = Path(__file__).resolve().parent
PROTOCOL = ROOT / "selector_calibration_protocol.json"


def previous_hashes():
    files = [p for p in ROOT.iterdir() if p.is_file()]
    files += [p for p in (ROOT / "runs").rglob("*") if p.is_file()]
    return {p.relative_to(ROOT).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}


def correctness_fixed(run, datasets, source):
    checks, oracles, f_exact = {}, {}, {}
    for name, data in datasets.items():
        ref = base.shapely.Polygon(base.VERTICES)
        oracle = set(map(int, data.ids[base.shapely.covers(ref, base.shapely.points(data.lon, data.lat))]))
        common = engine.prepare_common()
        result = {}
        for method in engine.METHODS:
            state, _ = engine.prepare_specific(method, data, common)
            ids, metrics = engine.query_function(method)(state, data)
            actual = set(map(int, ids))
            extra, missing, duplicate = actual - oracle, oracle - actual, len(ids) - len(actual)
            passed_probes = all((int(data.ids[i]) in actual) == truth for i, truth in enumerate(
                data.expected_probe_membership) if truth is not None)
            coverage = None if method == "direct" else base.check_coverage(state)
            result[method] = {"false_positive": len(extra), "false_negative": len(missing),
                              "duplicates": duplicate, "probes_passed": passed_probes,
                              "selected_points": len(ids), "shared_geometry_identity": state.region is common.region,
                              "coverage": coverage}
            if extra or missing or duplicate or not passed_probes or (coverage and not coverage["passed"]):
                base.dump_json(run / "correctness_failure.json", {"dataset": name, "method": method,
                               "extra": sorted(extra), "missing": sorted(missing), "check": result[method]})
                raise RuntimeError("统一入口正确性失败；保存诊断并停止性能结论")
            base.write_csv(run / "selected_ids" / name / f"{method}.csv",
                           [{"record_id": i} for i in sorted(actual)], ["record_id"])
            if method == "r7":
                f_exact[name] = metrics["exact_points"] / len(data.ids)
                # 校准特征与上一轮完整编码/边界类别交叉核对。
                old_counts = previous.read_csv(ROOT / "runs" / "bbox_20261111_001350_991724" / "bbox_candidate_counts.csv")
                old = next(r for r in old_counts if r["dataset"] == name and r["resolution"] == "7")
                assert metrics["exact_points"] == int(old["new_exact_points"])
        checks[name], oracles[name] = result, oracle
        print(f"{name}: 统一入口四方法ID/32个边界点通过，返回{len(oracle)}条。", flush=True)
    base.dump_json(run / "correctness_fixed.json", checks)
    base.dump_json(run / "offline_exact_features.json", {name: {"f_exact": f, "N": len(datasets[name].ids)}
                                                         for name, f in f_exact.items()})
    return checks, oracles, f_exact


def measure_fixed(method, data, k):
    query = engine.query_function(method)
    start = time.perf_counter_ns()
    common = engine.prepare_common()
    common_end = time.perf_counter_ns()
    state, detail = engine.prepare_specific(method, data, common)
    preparation_end = time.perf_counter_ns()
    for _ in range(k):
        ids, _ = query(state, data)
    end = time.perf_counter_ns()
    return ids, {"total_ns": end - start, "common_geometry_ns": common_end - start,
                 "specific_prepare_external_ns": preparation_end - common_end,
                 "continuous_queries_ns": end - preparation_end, "query_calls": k, **detail}


def benchmark_fixed(run, datasets, oracles, protocol):
    records = []
    cases = [(name, method, k) for name in protocol["datasets"] for method in engine.METHODS for k in protocol["K"]]
    rng = random.Random(protocol["order_seed"])
    was_enabled = gc.isenabled()
    gc.disable()
    try:
        for name in protocol["datasets"]:
            for method in engine.METHODS:
                measure_fixed(method, datasets[name], 1)
        for trial in range(1, protocol["fixed_method_repeats"] + 1):
            order = cases.copy()
            rng.shuffle(order)
            for position, (name, method, k) in enumerate(order, 1):
                ids, metrics = measure_fixed(method, datasets[name], k)
                actual = set(map(int, ids))
                if actual != oracles[name] or len(ids) != len(actual):
                    base.dump_json(run / "timing_correctness_failure.json", {"dataset": name, "method": method, "K": k,
                                   "extra": sorted(actual - oracles[name]), "missing": sorted(oracles[name] - actual)})
                    if records:
                        base.write_csv(run / "fixed_trials_before_failure.csv", records)
                    raise RuntimeError("固定方法计时后正确性失败")
                records.append({"trial": trial, "order_in_round": position, "dataset": name,
                                "method": method, "K": k, **metrics})
            print(f"统一入口固定方法交错计时：{trial}/{protocol['fixed_method_repeats']}", flush=True)
    finally:
        if was_enabled:
            gc.enable()
    return records


def fit_model(records, f_exact, protocol):
    features = np.asarray([f_exact[name] for name in protocol["datasets"]])
    x = np.column_stack((np.ones(len(features)), features))
    assert np.linalg.matrix_rank(x) == 2
    model = {"schema": 1, "f_min": float(features.min()), "f_max": float(features.max()),
             "sample_size": protocol["sample_size"], "sample_seed": protocol["sample_seed"], "methods": {}}
    training_rows, fit_rows, g_cases = [], [], []
    for name in protocol["datasets"]:
        g = [r["common_geometry_ns"] for r in records if r["dataset"] == name and r["K"] == 1000]
        g_cases.append(statistics.median(g) / 1e6)
    model["common_geometry_ms"] = statistics.median(g_cases)
    for method in engine.METHODS:
        costs, preparations = [], []
        for name in protocol["datasets"]:
            rows = [r for r in records if r["dataset"] == name and r["method"] == method and r["K"] == 1000]
            q = statistics.median(r["continuous_queries_ns"] / r["K"] for r in rows) / 1e6
            p = statistics.median(r["specific_prepare_external_ns"] for r in rows) / 1e6
            costs.append(q)
            preparations.append(p)
            training_rows.append({"dataset": name, "method": method, "f_exact": f_exact[name],
                                  "query_ms": q, "prepare_ms": p, "independent_case_observations": 1})
        coeff, _, rank, _ = np.linalg.lstsq(x, np.asarray(costs), rcond=None)
        a, b = map(float, coeff)
        assert rank == 2 and min(a + b * model["f_min"], a + b * model["f_max"]) > 0
        model["methods"][method] = {"prepare_ms": statistics.median(preparations), "a_ms": a, "b_ms": b,
                                    "mean_query_ms": statistics.mean(costs)}
        for name, actual in zip(protocol["datasets"], costs):
            predicted = a + b * f_exact[name]
            fit_rows.append({"dataset": name, "method": method, "actual_query_ms": actual,
                             "fitted_query_ms": predicted, "residual_ms": actual - predicted})
    return model, training_rows, fit_rows


def check_k100(records, model, f_exact, protocol):
    rows = []
    for name in protocol["datasets"]:
        for method in engine.METHODS:
            raw = [r for r in records if r["dataset"] == name and r["method"] == method and r["K"] == 100]
            params = model["methods"][method]
            predicted = model["common_geometry_ms"] + params["prepare_ms"] + 100 * (
                params["a_ms"] + params["b_ms"] * f_exact[name])
            measured = statistics.median(r["total_ns"] for r in raw) / 1e6
            relative = 100 * (predicted - measured) / measured
            rows.append({"dataset": name, "method": method, "K": 100, "offline_f_exact": f_exact[name],
                         "predicted_total_ms": predicted, "measured_total_ms": measured,
                         "relative_error_percent": relative, "absolute_error_percent": abs(relative),
                         "large_error_flag": abs(relative) > protocol["model"]["large_error_flag_percent"],
                         **{k: v for k, v in distribution_stats([r["total_ns"] for r in raw]).items() if k != "median_ms"}})
    return rows


def strategy_correctness(run, datasets, oracles, model, protocol):
    rows = []
    for name, data in datasets.items():
        raw = engine.RawCoordinates(data.lon, data.lat)
        for k in protocol["selector_cost_K"]:
            selections = {}
            for strategy in protocol["strategies"]:
                common = engine.prepare_common()
                choice = engine.choose(raw, common, k, model, strategy)
                selections[strategy] = choice
                state, _ = engine.prepare_specific(choice["selected"], data, common)
                ids, _ = engine.query_function(choice["selected"])(state, data)
                actual = set(map(int, ids))
                extra, missing, duplicates = actual - oracles[name], oracles[name] - actual, len(ids) - len(actual)
                probes_ok = all((int(data.ids[i]) in actual) == truth for i, truth in enumerate(
                    data.expected_probe_membership) if truth is not None)
                row = {"dataset": name, "K": k, "strategy": strategy, "selected": choice["selected"],
                       "base": choice["base"], "raw_argmin": choice["raw_argmin"],
                       "false_positive": len(extra), "false_negative": len(missing), "duplicates": duplicates,
                       "probes_passed": probes_ok, "common_geometry_identity": state.region is common.region,
                       "f_raw": choice["f_raw"], "f_used": choice["f_used"]}
                rows.append(row)
                if extra or missing or duplicates or not probes_ok:
                    base.dump_json(run / "strategy_failure.json", {"row": row, "extra": sorted(extra), "missing": sorted(missing)})
                    raise RuntimeError("选择策略正确性失败")
            a, b = selections["sample_two"], selections["sample_four"]
            assert a["sample_rows"] == b["sample_rows"] and a["costs"] == b["costs"] and a["base"] == b["base"]
            assert b["selected"] == a["selected"] or b["selected"] in ("r6", "r8")
    base.write_csv(run / "correctness_strategies.csv", rows)
    return rows


def benchmark_selection(datasets, model, protocol, diagnostic=False):
    records, traces = [], []
    repeat_count = protocol["selector_diagnostic_repeats"] if diagnostic else protocol["selector_normal_repeats"]
    cases = [(name, k, strategy) for name in protocol["datasets"] for k in protocol["selector_cost_K"]
             for strategy in protocol["strategies"]]
    raw = {name: engine.RawCoordinates(data.lon, data.lat) for name, data in datasets.items()}
    rng = random.Random(protocol["order_seed"] + (2 if diagnostic else 1))
    was_enabled = gc.isenabled()
    gc.disable()
    try:
        for name, k, strategy in cases:
            engine.choose(raw[name], engine.prepare_common(), k, model, strategy, diagnostic)
        for trial in range(1, repeat_count + 1):
            order = cases.copy()
            rng.shuffle(order)
            paired = {}
            for position, (name, k, strategy) in enumerate(order, 1):
                g_start = time.perf_counter_ns()
                common = engine.prepare_common()
                g_end = time.perf_counter_ns()
                s_start = time.perf_counter_ns()
                choice = engine.choose(raw[name], common, k, model, strategy, diagnostic)
                s_end = time.perf_counter_ns()
                record = {"trial": trial, "order_in_round": position, "dataset": name, "K": k,
                          "strategy": strategy, "common_geometry_ns": g_end - g_start,
                          "selection_ns": s_end - s_start, "selected": choice["selected"],
                          "base": choice["base"], "raw_argmin": choice["raw_argmin"],
                          "f_raw": choice["f_raw"], "f_used": choice["f_used"],
                          "sample_numerator": choice.get("sample_numerator"),
                          "unique_sample_cells": choice.get("unique_sample_cells", 0)}
                if diagnostic:
                    names = ("sampling_and_gather_ns", "r7_encoding_ns", "sample_dedup_ns",
                             "local_cell_classification_ns", "point_weighted_count_and_trace_ns",
                             "feature_total_ns", "prediction_ns", "final_decision_and_trace_ns")
                    record.update({key: choice.get("diagnostic_ns", {}).get(key, 0) for key in names})
                records.append(record)
                traces.append({"trial": trial, "dataset": name, "K": k, "strategy": strategy,
                               "sample_seed": model["sample_seed"], "diagnostic": diagnostic, "choice": choice})
                if strategy.startswith("sample_"):
                    paired.setdefault((name, k), {})[strategy] = choice
            for pair in paired.values():
                a, b = pair["sample_two"], pair["sample_four"]
                assert a["sample_rows"] == b["sample_rows"] and a["costs"] == b["costs"] and a["base"] == b["base"]
                assert b["selected"] == a["selected"] or b["selected"] in ("r6", "r8")
    finally:
        if was_enabled:
            gc.enable()
    return records, traces


def selection_summary(records):
    summary = []
    for name in dict.fromkeys(r["dataset"] for r in records):
        for k in (1, 100, 1000):
            for strategy in ("k_only", "sample_two", "sample_four"):
                rows = [r for r in records if r["dataset"] == name and r["K"] == k and r["strategy"] == strategy]
                summary.append({"dataset": name, "K": k, "strategy": strategy, "trials": len(rows),
                                **distribution_stats([r["selection_ns"] for r in rows]),
                                "common_geometry_median_ms": statistics.median(r["common_geometry_ns"] for r in rows) / 1e6,
                                "selected": rows[0]["selected"], "base": rows[0]["base"],
                                "raw_argmin": rows[0]["raw_argmin"], "f_raw": rows[0]["f_raw"], "f_used": rows[0]["f_used"]})
    return summary


def analyze_intervals(model):
    result, rows = {}, []
    for k in (100, 1000):
        analysis = decision_intervals(model, k)
        result[str(k)] = analysis
        for kind, collection in (("open", analysis["open_intervals"]), ("boundary", analysis["boundaries"])):
            for part in collection:
                rows.append({"K": k, "kind": kind, "left": part.get("left", part.get("f")),
                             "right": part.get("right", part.get("f")), "sample_two": part["sample2"]["selected"],
                             "sample_four": part["sample4"]["selected"], "raw_argmin": part["raw_argmin"],
                             "base": part["sample4"]["base"], "extra_possible": part["extra_possible"]})
        for side, part in analysis["clipped_tails"].items():
            rows.append({"K": k, "kind": "clip_" + side, "left": 0 if side == "left" else model["f_max"],
                         "right": model["f_min"] if side == "left" else 1,
                         "sample_two": part["sample2"]["selected"], "sample_four": part["sample4"]["selected"],
                         "raw_argmin": part["raw_argmin"], "base": part["sample4"]["base"],
                         "extra_possible": part["extra_possible"]})
    has_extra = any(row["sample_four"] in ("r6", "r8") for row in rows)
    assert all(row["sample_two"] == row["sample_four"] or row["sample_four"] in ("r6", "r8") for row in rows)
    return result, rows, has_extra


def write_results(run, model, k100, selections, intervals, has_extra, cost_ledger):
    lines = ["# 统一准备与轻量选择：旧数据校准", "",
             "仅使用旧四组模拟数据。未生成或查看六组独立验证数据。查询函数保持冻结，原版和旧记录保留。", "",
             "## 校准参数", "",
             "预计总成本 = 共同几何G + 方法专属准备P + K × (a + b × f)。单位毫秒，f使用0到1的小数。",
             "含义：P是该方法尚需的固定准备费用，a+b×f是正常连续查询的单次成本估计；每方法只有两个查询参数。", "",
             f"共同几何估计G={model['common_geometry_ms']:.6f}毫秒；比例范围"
             f"[{model['f_min']:.9f}, {model['f_max']:.9f}]。", "",
             "| 方法 | P：毫秒 | a：毫秒/次 | b：毫秒/次/单位比例 | K-only单次均值：毫秒 |",
             "| --- | ---: | ---: | ---: | ---: |"]
    for method, p in model["methods"].items():
        lines.append(f"| {method} | {p['prepare_ms']:.8f} | {p['a_ms']:.9f} | {p['b_ms']:.9f} | {p['mean_query_ms']:.9f} |")
    lines += ["", "各旧数据组在K=1000的7次正常连续查询段除以1000后取中位数，每组只贡献一个观测；4组等权最小二乘拟合直线。P取四组专属准备中位数的中位数。共同G独立统计。",
              "K=100完全不参与系数拟合，仅检查模型在较短执行长度下的适用性。以下使用旧数据的离线准确f，不是线上样本，也不是独立验证。", "",
              "## K=100预测检查", "",
              "相对误差 = (预测总成本−实测总成本)/实测总成本；正数表示高估，负数表示低估。", "",
              "| 数据 | 方法 | 预测总计ms | 实测总计ms | 相对误差% |",
              "| --- | --- | ---: | ---: | ---: |"]
    for r in k100:
        lines.append(f"| {r['dataset']} | {r['method']} | {r['predicted_total_ms']:.4f} | "
                     f"{r['measured_total_ms']:.4f} | {r['relative_error_percent']:+.2f} |")
    lines += ["", "## 正常选择过程费用", "",
              "费用列只含选择S，共同几何G另列；不含选中方法准备和查询。每策略独立执行，不借用另一策略结果。",
              "重复使用相同种子只反映计时波动，不形成多批独立样本。K=1直接返回直接查询，不抽样。", "",
              "| 数据 | K | 策略 | S中位数ms | S第25–75百分位ms | 另计G中位数ms | 选择 | 原始f/预测f |",
              "| --- | ---: | --- | ---: | --- | ---: | --- | --- |"]
    for r in selections:
        feature = "—" if r["f_raw"] is None else f"{r['f_raw']:.4f}/{r['f_used']:.4f}"
        lines.append(f"| {r['dataset']} | {r['K']} | {r['strategy']} | {r['median_ms']:.6f} | "
                     f"{r['q25_ms']:.6f}–{r['q75_ms']:.6f} | {r['common_geometry_median_ms']:.6f} | {r['selected']} | {feature} |")
    lines += ["", "## 解析决策区间", "",
              "先在direct/r7中确定同一个基准B，再检查新增档是否严格便宜超过2%；未过门槛四选一保持B。",
              "raw_argmin仅记录不保守时的四线最低方法，不增加一套计时策略。开区间与交点边界分别保存在CSV/JSON，阈值等式不允许新增切换。", "",
              "| K | 类型 | 左端 | 右端 | 样本二选一 | 样本四选一 | 不保守最低 |",
              "| ---: | --- | ---: | ---: | --- | --- | --- |"]
    for r in intervals:
        lines.append(f"| {r['K']} | {r['kind']} | {r['left']:.9f} | {r['right']:.9f} | "
                     f"{r['sample_two']} | {r['sample_four']} | {r['raw_argmin']} |")
    lines += ["", ("存在最终选择新增档的区间，但本轮没有独立验证，不能声称有收益。" if has_extra else
                     "**校准范围及裁剪尾部均不存在最终选择r6/r8的区间。当前规则退化为样本二选一，停止在本轮讨论，不修改阈值或数据制造新增选择。**"), "",
              "## 正确性与费用边界", "",
              "四固定方法和三种策略在旧数据上逐条核对ID，误选、漏选、重复均为0；32个测试点通过。共同几何按对象身份真正复用，只创建/预处理一次。样本编码和局部格轮廓不传给全量准备，不预扣准备费用。",
              "固定方法总时间是G+P+连续K次查询的一次外层实测；正常S独立实测，分项诊断仅作解释。本轮没有把S与几个中位数相加宣称策略端到端收益。", "",
              "## 离线成本", "",
              "以下为本轮校准工作的墙钟阶段费用；核心计时是其中子集，不将二者再相加。全部单位秒。", "",
              "| 阶段 | 墙钟秒 |", "| --- | ---: |"]
    for name, ns in cost_ledger.items():
        lines.append(f"| {name} | {ns / 1e9:.6f} |")
    lines += ["", "源码、协议、输入指纹、online_model.json、离线训练明细、K100检查、正常/诊断选择原始记录、所有交点方程和边界选择均在本目录。",
              "后续六组数据的预留方案保持不变，尚未生成；现有四组全部属于校准/探索，不是最终验证。", ""]
    (run / "校准结果.md").write_text("\n".join(lines), encoding="utf-8")


def main():
    job_start = time.perf_counter_ns()
    protocol_bytes = PROTOCOL.read_bytes()
    protocol = json.loads(protocol_bytes)
    protected = previous_hashes()
    run = ROOT / "runs" / datetime.now().astimezone().strftime("selector_cal_%Y%m%d_%H%M%S_%f")
    run.mkdir(parents=True)
    (run / "protocol.json").write_bytes(protocol_bytes)
    base.dump_json(run / "protocol_record.json", {"recorded_before_timing": datetime.now().astimezone().isoformat(),
                                                 "sha256": hashlib.sha256(protocol_bytes).hexdigest()})
    base.dump_json(run / "previous_files_sha256.json", protected)
    for name in ("selector_engine.py", "selector_decision.py", "calibrate_selector.py", "test_selector_engine.py",
                 "test_selector_decision.py", "run_experiment.py", "bbox_ablation.py", "requirements.txt"):
        path = run / "source_snapshots" / name
        path.parent.mkdir(exist_ok=True)
        path.write_bytes((ROOT / name).read_bytes())
    env = base.environment_info()
    env["method_order_seed"] = protocol["order_seed"]
    base.dump_json(run / "environment.json", env)
    print(f"统一入口旧数据校准目录：{run}", flush=True)
    ledger, datasets, inputs = {}, {}, {}
    source = ROOT / protocol["source_run"]
    t = time.perf_counter_ns()
    for name in protocol["datasets"]:
        folder = source / "datasets" / name
        datasets[name] = previous.load_saved_data(folder)
        for filename in ("simulated_points.csv", "simulated_sea.geojson"):
            src = folder / "data" / filename
            dst = run / "input_data" / name / filename
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            digest = hashlib.sha256(src.read_bytes()).hexdigest()
            assert digest == hashlib.sha256(dst.read_bytes()).hexdigest()
            inputs[dst.relative_to(run).as_posix()] = {"source": src.relative_to(ROOT).as_posix(), "sha256": digest}
    base.dump_json(run / "input_manifest.json", inputs)
    checks, oracles, f_exact = correctness_fixed(run, datasets, source)
    ledger["old_input_loading_and_correctness_wall_ns"] = time.perf_counter_ns() - t
    t = time.perf_counter_ns()
    fixed = benchmark_fixed(run, datasets, oracles, protocol)
    ledger["fixed_calibration_benchmark_wall_ns"] = time.perf_counter_ns() - t
    base.write_csv(run / "fixed_method_trials.csv", fixed)
    fixed_summary = []
    for name in protocol["datasets"]:
        for k in protocol["K"]:
            for method in engine.METHODS:
                rows = [r for r in fixed if r["dataset"] == name and r["K"] == k and r["method"] == method]
                fixed_summary.append({"dataset": name, "K": k, "method": method, "trials": len(rows),
                                      **distribution_stats([r["total_ns"] for r in rows])})
    base.write_csv(run / "fixed_method_summary.csv", fixed_summary)
    t = time.perf_counter_ns()
    model, training, residuals = fit_model(fixed, f_exact, protocol)
    k100 = check_k100(fixed, model, f_exact, protocol)
    ledger["fit_and_K100_check_wall_ns"] = time.perf_counter_ns() - t
    base.dump_json(run / "online_model.json", model)
    base.write_csv(run / "offline_training_cases.csv", training)
    base.write_csv(run / "fit_residuals.csv", residuals)
    base.write_csv(run / "K100_prediction_check.csv", k100)
    base.write_csv(run / "coefficients.csv", [{"method": m, **p} for m, p in model["methods"].items()])
    print("K1000四组聚合拟合完成；K100仅作预测核对。开始正常选择费用测量。", flush=True)
    t = time.perf_counter_ns()
    normal, traces = benchmark_selection(datasets, model, protocol)
    ledger["normal_selection_cost_benchmark_wall_ns"] = time.perf_counter_ns() - t
    base.write_csv(run / "selection_normal_trials.csv", normal)
    with (run / "selection_normal_traces.jsonl").open("w", encoding="utf-8") as stream:
        for trace in traces:
            stream.write(json.dumps(trace, ensure_ascii=False) + "\n")
    selection_stats = selection_summary(normal)
    base.write_csv(run / "selection_normal_summary.csv", selection_stats)
    t = time.perf_counter_ns()
    diagnostic, diagnostic_traces = benchmark_selection(datasets, model, protocol, diagnostic=True)
    ledger["diagnostic_selection_benchmark_wall_ns"] = time.perf_counter_ns() - t
    base.write_csv(run / "selection_diagnostic_trials.csv", diagnostic)
    t = time.perf_counter_ns()
    strategy_checks = strategy_correctness(run, datasets, oracles, model, protocol)
    ledger["strategy_correctness_wall_ns"] = time.perf_counter_ns() - t
    t = time.perf_counter_ns()
    detailed_intervals, interval_rows, has_extra = analyze_intervals(model)
    ledger["analytic_decision_intervals_wall_ns"] = time.perf_counter_ns() - t
    base.dump_json(run / "decision_intervals.json", detailed_intervals)
    base.write_csv(run / "decision_intervals.csv", interval_rows)
    base.dump_json(run / "summary.json", {"calibration_only": True, "independent_validation_generated": False,
                                         "model": model, "K100_checks": k100, "selection_costs": selection_stats,
                                         "decision_intervals": interval_rows, "new_methods_have_selection_region": has_extra,
                                         "fixed_correctness": checks, "strategy_correctness": strategy_checks})
    core_costs = {"fixed_method_core_measured_ns": sum(r["total_ns"] for r in fixed),
                  "normal_selection_core_measured_ns": sum(r["selection_ns"] for r in normal),
                  "normal_selection_common_geometry_ns": sum(r["common_geometry_ns"] for r in normal),
                  "diagnostic_selection_core_measured_ns": sum(r["selection_ns"] for r in diagnostic)}
    write_results(run, model, k100, selection_stats, interval_rows, has_extra, ledger)
    ledger["whole_job_wall_until_accounting_export_ns"] = time.perf_counter_ns() - job_start
    base.dump_json(run / "offline_costs.json", {"wall_stages_ns": ledger, "core_subsets_ns": core_costs,
                                               "note": "overall wall encloses stages and core subsets; do not add them again"})
    assert PROTOCOL.read_bytes() == protocol_bytes
    assert all(hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == digest for name, digest in protected.items())
    base.dump_json(run / "status.json", {"status": "completed_calibration_only", "new_methods_have_region": has_extra,
                                        "stop_reason": "discuss_degenerate_rule" if not has_extra else "await_independent_validation_authorization",
                                        "fixed_trials": len(fixed), "normal_selection_trials": len(normal),
                                        "diagnostic_trials": len(diagnostic), "old_files_unchanged": True,
                                        "new_validation_data_generated_or_read": False})
    base.dump_json(run / "sha256_manifest.json", {p.relative_to(run).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
                                                 for p in run.rglob("*") if p.is_file()})
    (ROOT / "LATEST_SELECTOR_CALIBRATION.txt").write_text(run.relative_to(ROOT).as_posix() + "\n", encoding="utf-8")
    for method, params in model["methods"].items():
        print(method, params, flush=True)
    print(f"K100绝对相对误差：中位数{statistics.median(r['absolute_error_percent'] for r in k100):.2f}%，"
          f"最大{max(r['absolute_error_percent'] for r in k100):.2f}%。", flush=True)
    print(f"最终新增档有选择区间：{has_extra}。未进入独立验证。", flush=True)


if __name__ == "__main__":
    main()
