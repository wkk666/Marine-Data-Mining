"""只重算保存记录与公式，不执行性能测量或生成验证数据。"""
import hashlib
import json
from pathlib import Path
import statistics

import calibrate_selector as cal
import bbox_ablation as previous
from selector_decision import predict_costs, decide

ROOT = Path(__file__).resolve().parent


def main():
    run = ROOT / (ROOT / "LATEST_SELECTOR_CALIBRATION.txt").read_text("utf-8").strip()
    read = previous.read_csv
    protocol = json.loads((run / "protocol.json").read_text("utf-8"))
    model = json.loads((run / "online_model.json").read_text("utf-8"))
    raw = read(run / "fixed_method_trials.csv")
    converted = [{key: int(value) if key.endswith("_ns") or key in ("K", "trial", "order_in_round", "query_calls") else value
                  for key, value in row.items()} for row in raw]
    assert len(raw) == 224
    for name in protocol["datasets"]:
        for method in protocol["methods"]:
            for k in (100, 1000):
                rows = [r for r in converted if r["dataset"] == name and r["method"] == method and r["K"] == k]
                assert len(rows) == 7
                for r in rows:
                    assert r["query_calls"] == k
                    assert r["total_ns"] == r["common_geometry_ns"] + r["specific_prepare_external_ns"] + r["continuous_queries_ns"]
    f_records = json.loads((run / "offline_exact_features.json").read_text("utf-8"))
    f_exact = {name: row["f_exact"] for name, row in f_records.items()}
    refit, _, _ = cal.fit_model(converted, f_exact, protocol)
    assert refit == model
    # K100发生任何变化均不能影响拟合。只在内存中改检查记录，不改保存文件。
    altered = [dict(r, continuous_queries_ns=r["continuous_queries_ns"] * 10) if r["K"] == 100 else r for r in converted]
    assert cal.fit_model(altered, f_exact, protocol)[0] == model
    expected_k100 = cal.check_k100(converted, model, f_exact, protocol)
    saved_k100 = read(run / "K100_prediction_check.csv")
    for expected, saved in zip(expected_k100, saved_k100, strict=True):
        for key in ("predicted_total_ms", "measured_total_ms", "relative_error_percent"):
            assert abs(expected[key] - float(saved[key])) < 1e-12
    selection_rows = read(run / "selection_normal_trials.csv")
    saved_summary = read(run / "selection_normal_summary.csv")
    assert len(selection_rows) == 1116 and len(read(run / "selection_diagnostic_trials.csv")) == 612
    for s in saved_summary:
        rows = [r for r in selection_rows if r["dataset"] == s["dataset"] and r["K"] == s["K"] and r["strategy"] == s["strategy"]]
        assert len(rows) == 31
        stats = cal.distribution_stats([int(r["selection_ns"]) for r in rows])
        assert all(abs(float(s[key]) - value) < 1e-12 for key, value in stats.items())
    # 仅用已记录的样本比例重算预测；不重新抽样、不执行新的查询或计时。
    sample_checks = []
    for r in saved_k100:
        sample = next(s for s in saved_summary if s["dataset"] == r["dataset"] and s["K"] == "100"
                      and s["strategy"] == "sample_two")
        f_raw = float(sample["f_raw"])
        prediction = model["common_geometry_ms"] + predict_costs(model, f_raw, 100)[r["method"]]
        measured = float(r["measured_total_ms"])
        sample_checks.append({"dataset": r["dataset"], "method": r["method"], "K": 100,
                              "sample_f_raw": f_raw, "sample_f_used": float(sample["f_used"]),
                              "predicted_method_total_ms": prediction, "measured_fixed_method_total_ms": measured,
                              "relative_error_percent": 100 * (prediction - measured) / measured,
                              "note": "fixed-method cost prediction only; selection fee S excluded; not end-to-end selector measurement"})
    traces = [json.loads(line) for line in (run / "selection_normal_traces.jsonl").read_text("utf-8").splitlines()]
    grouped = {}
    for t in traces:
        choice = t["choice"]
        if t["K"] == 1:
            assert choice["selected"] == "direct" and choice["sample_rows"] == []
            continue
        if t["strategy"].startswith("sample_"):
            assert len(choice["sample_rows"]) == len(set(choice["sample_rows"])) == 200
            costs = predict_costs(model, choice["f_raw"], t["K"])
            assert costs == choice["costs"]
            decision = decide(costs, t["strategy"] == "sample_four")
            assert decision["selected"] == choice["selected"] and decision["base"] == choice["base"]
            grouped.setdefault((t["trial"], t["dataset"], t["K"]), {})[t["strategy"]] = choice
    for pair in grouped.values():
        a, b = pair["sample_two"], pair["sample_four"]
        assert a["costs"] == b["costs"] and a["sample_rows"] == b["sample_rows"] and a["base"] == b["base"]
        assert a["selected"] == b["selected"] or b["selected"] in ("r6", "r8")
    detailed, interval_rows, has_extra = cal.analyze_intervals(model)
    assert detailed == json.loads((run / "decision_intervals.json").read_text("utf-8"))
    discrete_extra = {}
    for k in (100, 1000):
        discrete_extra[str(k)] = [i / 200 for i in range(201)
                                  if decide(predict_costs(model, i / 200, k), True)["selected"] in ("r6", "r8")]
    correctness = read(run / "correctness_strategies.csv")
    assert len(correctness) == 36
    assert all(r["false_positive"] == r["false_negative"] == r["duplicates"] == "0"
               and r["probes_passed"] == r["common_geometry_identity"] == "True" for r in correctness)
    fixed_checks = json.loads((run / "correctness_fixed.json").read_text("utf-8"))
    assert all(c["false_positive"] == c["false_negative"] == c["duplicates"] == 0
               and c["probes_passed"] and c["shared_geometry_identity"] for checks in fixed_checks.values() for c in checks.values())
    inputs = json.loads((run / "input_manifest.json").read_text("utf-8"))
    assert all((run / path).read_bytes() == (ROOT / value["source"]).read_bytes() for path, value in inputs.items())
    old = json.loads((run / "previous_files_sha256.json").read_text("utf-8"))
    assert all(hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == digest for name, digest in old.items())
    manifest = json.loads((run / "sha256_manifest.json").read_text("utf-8"))
    assert all(hashlib.sha256((run / name).read_bytes()).hexdigest() == digest for name, digest in manifest.items())
    audit = {"status": "passed", "fixed_normal_trials": 224, "normal_selection_trials": 1116,
             "diagnostic_selection_trials": 612, "coefficients_refitted_exactly": True,
             "K100_excluded_from_fit_confirmed": True, "all_prediction_checks_recomputed": True,
             "all_normal_selection_cost_medians_recomputed": True, "analytic_intervals_recomputed": True,
             "possible_sample_grid_extra_selections": discrete_extra, "old_files_verified": len(old),
             "new_validation_data_generated_or_read": False}
    cal.base.write_csv(run / "K100_sample_prediction_check.csv", sample_checks)
    cal.base.dump_json(run / "artifact_validation.json", audit)
    manifest["artifact_validation.json"] = hashlib.sha256((run / "artifact_validation.json").read_bytes()).hexdigest()
    manifest["K100_sample_prediction_check.csv"] = hashlib.sha256((run / "K100_sample_prediction_check.csv").read_bytes()).hexdigest()
    cal.base.dump_json(run / "sha256_manifest.json", manifest)
    print(json.dumps(audit, ensure_ascii=False))


if __name__ == "__main__":
    main()
