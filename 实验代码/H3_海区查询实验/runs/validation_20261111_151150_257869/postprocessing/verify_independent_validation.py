"""只读复核已完成的六组独立验证；不导入实验模块、不生成点、不运行查询或计时。

python verify_independent_validation.py
python verify_independent_validation.py --write-audit

默认仅打印审计结果。--write-audit 只新增审计JSON和本脚本副本；拒绝覆盖已有
文件，不改原始 sha256_manifest.json。后续交付清单由主任务另行维护。
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import csv
from datetime import datetime
import hashlib
import json
import math
from pathlib import Path
import random
import statistics
import sys

RUN_NAME = "validation_20261111_151150_257869"
ROOT = Path(__file__).resolve().parent
if ROOT.name == "postprocessing" and ROOT.parent.name == RUN_NAME and ROOT.parents[1].name == "runs":
    ROOT = ROOT.parents[2]
RUN = ROOT / "runs" / RUN_NAME
POLICIES = ("k_only", "sample_two", "sample_four")
FIXED = {"fixed_direct": "direct", "fixed_r7": "r7", "fixed_r6": "r6", "fixed_r8": "r8"}
METHODS = ("direct", "r6", "r7", "r8")
CHOICE_KEYS = ("selected", "base", "raw_argmin", "f_raw", "f_used", "sample_rows",
               "sample_numerator", "sample_size", "costs", "extra_qualified")
EPS = 1e-12


def require(condition, message):
    if not condition:
        raise AssertionError(message)


def same(actual, expected, message):
    if expected is None:
        require(actual is None or actual == "", f"{message}: expected empty, got {actual!r}")
    elif isinstance(expected, bool):
        require(actual == expected or actual == str(expected), f"{message}: {actual!r} != {expected!r}")
    elif isinstance(expected, (int, float)):
        try:
            value = float(actual)
        except (TypeError, ValueError) as exc:
            raise AssertionError(f"{message}: invalid number {actual!r}") from exc
        require(math.isfinite(value) and math.isclose(value, expected, rel_tol=1e-9, abs_tol=1e-9),
                f"{message}: {actual!r} != {expected!r}")
    else:
        require(actual == expected, f"{message}: {actual!r} != {expected!r}")


def json_file(path):
    return json.loads(path.read_text(encoding="utf-8"))


def csv_file(name):
    with (RUN / name).open(encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream))


def json_lines(name):
    with (RUN / name).open(encoding="utf-8") as stream:
        return [json.loads(line) for line in stream if line.strip()]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def within(root, relative):
    path = (root / relative).resolve()
    require(path.is_relative_to(root.resolve()), f"path escapes intended directory: {relative}")
    return path


def verify_manifest(root, manifest):
    for name, expected in manifest.items():
        path = within(root, name)
        require(path.is_file(), f"missing hash target: {path}")
        same(digest(path), expected, f"SHA-256 {path}")
    return len(manifest)


def percentile(values, p):
    ordered = sorted(values)
    index = (len(ordered) - 1) * p
    lo, hi = math.floor(index), math.ceil(index)
    return ordered[lo] + (ordered[hi] - ordered[lo]) * (index - lo)


def stats(values):
    middle = statistics.median(values)
    return {"median_ms": middle, "q25_ms": percentile(values, .25),
            "q75_ms": percentile(values, .75), "min_ms": min(values), "max_ms": max(values),
            "mad_ms": statistics.median(abs(x - middle) for x in values)}


def unique_index(rows, fields, expected_count, label):
    result = {tuple(row[field] for field in fields): row for row in rows}
    require(len(result) == len(rows) == expected_count, f"{label}: missing or duplicate keys")
    return result


def scores(model, feature, k):
    # 独立复核冻结公式：准确 f 的反事实也裁剪，不做重新拟合或额外截断。
    f = min(model["f_max"], max(model["f_min"], feature))
    return {m: model["methods"][m]["prepare_ms"] + k * model["methods"][m]["a_ms"]
            + k * model["methods"][m]["b_ms"] * f for m in METHODS}


def argmin_ties(values, priority):
    chosen = priority[0]
    for method in priority[1:]:
        if values[method] - values[chosen] < -EPS:
            chosen = method
    return chosen


def decision(values, allow_extra):
    base = "direct" if values["direct"] - 1.02 * values["r7"] <= EPS else "r7"
    qualified = [m for m in ("r6", "r8") if values[m] - .98 * values[base] < -EPS]
    chosen = argmin_ties(values, qualified) if allow_extra and qualified else base
    return {"base": base, "selected": chosen, "extra_qualified": qualified,
            "raw_argmin": argmin_ties(values, ("direct", "r7", "r6", "r8"))}


def audit():
    notes = []
    counts = {}
    original_manifest_bytes = (RUN / "sha256_manifest.json").read_bytes()
    manifest = json.loads(original_manifest_bytes)
    counts["original_manifest_files"] = verify_manifest(RUN, manifest)
    counts["previous_files"] = verify_manifest(ROOT, json_file(RUN / "previous_files_sha256.json"))
    freeze = json_file(RUN / "freeze_manifest.json")
    protocol = json_file(RUN / "protocol.json")
    model = json_file(RUN / "online_model.json")
    same(digest(RUN / "protocol.json"), freeze["protocol_sha256"], "frozen protocol")
    same(digest(ROOT / "independent_validation_protocol.json"), freeze["protocol_sha256"], "source protocol")
    same(digest(RUN / "online_model.json"), protocol["model_sha256"], "model protocol hash")
    same(digest(RUN / "online_model.json"), freeze["model_sha256"], "model freeze hash")
    same((RUN / "online_model.json").read_bytes(), within(ROOT, protocol["model_source"]).read_bytes(),
         "full precision original model bytes")
    for name, expected in freeze["code_sha256"].items():
        same(digest(within(RUN / "snapshots", name)), expected, f"snapshot {name}")
        same(digest(within(ROOT, name)), expected, f"unchanged current code {name}")
    for name, record in json_file(RUN / "frozen_code_provenance.json").items():
        same(digest(within(ROOT, record["source"])), record["sha256"], f"calibration source {name}")
        same(digest(within(RUN / "snapshots", name)), record["sha256"], f"calibration provenance {name}")
    notes.append("Original output manifest, old files, full-precision model, protocol and source provenance match.")

    cases = {case["name"]: case for case in protocol["datasets"]}
    require(len(cases) == 6 and protocol["K"] == [1, 100, 1000] and protocol["repeats"] == 7,
            "frozen six-dataset/three-K/seven-repeat protocol")
    events = json_lines("events.jsonl")
    require(all(a["monotonic_ns"] < b["monotonic_ns"] for a, b in zip(events, events[1:])),
            "events must be strictly ordered by monotonic clock")
    names = [event["event"] for event in events]
    stages = ("code_model_protocol_frozen", "all_recommendations_locked", "correctness_preflight_started",
              "correctness_preflight_completed", "uniform_warmup_completed", "formal_timing_started",
              "formal_timing_completed", "posthoc_diagnostics_started", "validation_completed_no_refit")
    require(all(names.count(name) == 1 for name in stages), "one event per required phase")
    positions = [names.index(name) for name in stages]
    require(positions == sorted(positions), "required phase order")
    generated = [e for e in events if e["event"] == "dataset_generated"]
    require(len(generated) == 6 and {e["dataset"] for e in generated} == set(cases), "six generated datasets only")
    require(all(positions[0] < events.index(e) < positions[1] for e in generated),
            "all generation must occur after freeze and before all recommendations lock")
    rounds = [e for e in events if e["event"] == "formal_round_completed"]
    require([e["trial"] for e in rounds] == list(range(1, 8)), "seven formal round events")
    require(all(e["records"] == e["trial"] * 126 for e in rounds), "126 tasks per formal round")
    lock = json_file(RUN / "prerecommendations_lock.json")
    same(lock["count"], 54, "prerecommendation lock count")
    same(lock["all_saved_before_full_correctness_warmup_or_timing"], True, "prerecommendation lock flag")
    same(digest(RUN / "prerecommendations.jsonl"), lock["sha256"], "recommendations immutable hash")
    lock_event = events[names.index("all_recommendations_locked")]
    same(lock_event["sha256"], lock["sha256"], "recommendation event hash")
    same(lock_event["count"], 54, "recommendation event count")
    notes.append("All 54 recommendations locked before full correctness, warmup, timing and posthoc diagnostics.")

    correctness = json_file(RUN / "correctness.json")
    data_hashes = json_file(RUN / "data_hashes.json")
    dataset_rows, selected_sets = {}, {}
    for name, case in cases.items():
        points_path = RUN / "datasets" / name / "data" / "simulated_points.csv"
        same(digest(points_path), data_hashes[name], f"data hash {name}")
        generated_event = next(e for e in generated if e["dataset"] == name)
        same(generated_event["seed"], case["seed"], f"data seed {name}")
        same(generated_event["input_sha256"], data_hashes[name], f"event data hash {name}")
        rows = csv_file(f"datasets/{name}/data/simulated_points.csv")
        require(len(rows) == len({int(r["record_id"]) for r in rows}) == 10032, f"unique input IDs {name}")
        dataset_rows[name] = rows
        categories = Counter(r["category"] for r in rows)
        if case["distribution"] == "uniform":
            same(categories["uniform_simulated"], 10000, f"uniform mixture count {name}")
        else:
            for category, field in (("uniform_component", "uniform_points"),
                                    ("boundary_band_inside", "inside_band"),
                                    ("boundary_band_outside", "outside_band")):
                same(categories[category], case[field], f"mixture count {name}/{category}")
        probes = [r for r in rows if r["probe_expected_inside"] != ""]
        require(len(probes) == 32 and sum(int(r["probe_expected_inside"]) for r in probes) == 24,
                f"independent boundary expectations {name}")
        selections = []
        for method in METHODS:
            entries = csv_file(f"selected_ids/{name}/{method}.csv")
            ids = {int(r["record_id"]) for r in entries}
            require(len(ids) == len(entries) and ids <= {int(r["record_id"]) for r in rows},
                    f"unique selected IDs within input {name}/{method}")
            require(all((int(r["record_id"]) in ids) == bool(int(r["probe_expected_inside"])) for r in probes),
                    f"boundary membership {name}/{method}")
            check = correctness[name][method]
            for field in ("false_positive", "false_negative", "duplicates"):
                same(check[field], 0, f"saved correctness {name}/{method}/{field}")
            same(check["boundary_probes_passed"], True, f"saved boundary status {name}/{method}")
            same(check["selected_points"], len(ids), f"selected count {name}/{method}")
            if method != "direct":
                same(check["coverage"]["passed"], True, f"coverage audit {name}/{method}")
                same(check["coverage"]["uncovered_area_square_degrees"], 0, f"uncovered area {name}/{method}")
            selections.append(ids)
        require(all(ids == selections[0] for ids in selections), f"all fixed ID sets equal {name}")
        selected_sets[name] = selections[0]
    counts["independent_datasets"] = 6
    counts["boundary_probes_checked"] = 6 * 32
    notes.append("Saved fixed ID sets agree; all 192 boundary probes and saved full-cover checks pass.")

    prerecs = json_lines("prerecommendations.jsonl")
    pre = unique_index(prerecs, ("dataset", "K", "policy"), 54, "prerecommendations")
    require(set(pre) == {(name, k, policy) for name in cases for k in protocol["K"] for policy in POLICIES},
            "complete prerecommendation scenario keys")
    for (name, k, policy), record in pre.items():
        choice = record["choice"]
        same(record["sample_seed"], 73001, "frozen sampling seed")
        if k == 1:
            same(choice["selected"], "direct", "K1 shortcut")
            require(choice["sample_rows"] == [] and choice["costs"] == {} and choice["f_raw"] is None,
                    "K1 shortcut must not sample or predict")
            continue
        if policy == "k_only":
            require(choice["sample_rows"] == [] and choice["f_raw"] is None, "K-only must not sample")
            expected_costs = {m: model["methods"][m]["prepare_ms"]
                              + k * model["methods"][m]["mean_query_ms"] for m in METHODS}
        else:
            indices = choice["sample_rows"]
            require(len(indices) == len(set(indices)) == 200 and all(0 <= i < 10032 for i in indices),
                    "200 valid distinct sample row positions")
            same(record["sample_record_ids"], [int(dataset_rows[name][i]["record_id"]) for i in indices],
                 "sample row positions map to saved record IDs")
            same(choice["f_raw"], choice["sample_numerator"] / 200, "point-weighted sample fraction")
            same(choice["f_used"], min(model["f_max"], max(model["f_min"], choice["f_raw"])), "sample clipping")
            expected_costs = scores(model, choice["f_raw"], k)
        for method in METHODS:
            same(choice["costs"][method], expected_costs[method], f"frozen cost {name}/{k}/{policy}/{method}")
        for key, expected in decision(expected_costs, policy == "sample_four").items():
            same(choice[key], expected, f"frozen two-percent decision {name}/{k}/{policy}/{key}")
    for name in cases:
        for k in protocol["K"]:
            a, b = pre[name, k, "sample_two"]["choice"], pre[name, k, "sample_four"]["choice"]
            for key in ("sample_rows", "sample_numerator", "f_raw", "f_used", "costs", "base"):
                same(a[key], b[key], f"two/four shared input {name}/{k}/{key}")
    counts["prerecommendations"] = 54

    raw = csv_file("formal_trials.csv")
    index = {}
    groups = defaultdict(list)
    for row in raw:
        name, k, scheme, trial = row["dataset"], int(row["K"]), row["scheme"], int(row["trial"])
        key = name, k, scheme, trial
        require(key not in index and name in cases and k in protocol["K"] and scheme in protocol["executions"],
                "unique valid formal key")
        require(int(row["total_ns"]) > 0, "positive outer wall interval")
        for field in ("false_positive", "false_negative", "duplicates"):
            same(row[field], 0, f"formal recorded correctness {key}/{field}")
        same(row["prerecommendation_match"], True, f"formal pre-choice flag {key}")
        expected_method = FIXED[scheme] if scheme in FIXED else pre[name, k, scheme]["choice"]["selected"]
        same(row["chosen_method"], expected_method, f"chosen method {key}")
        same(row["selected_records"], len(selected_sets[name]), f"formal selected count {key}")
        index[key] = row
        groups[name, k, scheme].append(row)
    require(len(raw) == 882 and len(groups) == 126, "882 formal records in 126 groups")
    require(all(len(v) == 7 and {int(r["trial"]) for r in v} == set(range(1, 8)) for v in groups.values()),
            "seven unique trial numbers per group")
    ordered_cases = [(case["name"], k, scheme) for case in protocol["datasets"]
                     for k in protocol["K"] for scheme in protocol["executions"]]
    rng = random.Random(protocol["order_seed"])
    for trial in range(1, 8):
        order = ordered_cases.copy()
        rng.shuffle(order)
        for position, (name, k, scheme) in enumerate(order, 1):
            same(index[name, k, scheme, trial]["order_in_round"], position, "frozen randomized round order")
    traces = unique_index(json_lines("formal_choice_traces.jsonl"), ("dataset", "K", "scheme", "trial"),
                          882, "formal choice traces")
    require(set(traces) == set(index), "formal trace key set")
    for (name, k, scheme, trial), trace in traces.items():
        choice = trace["choice"]
        if scheme in FIXED:
            same(choice["selected"], FIXED[scheme], "fixed trace method")
        else:
            saved = pre[name, k, scheme]["choice"]
            require(all(choice.get(key) == saved.get(key) for key in CHOICE_KEYS),
                    f"immutable full prerecommendation trace {name}/{k}/{scheme}/{trial}")
    counts.update(formal_records=882, formal_groups=126, formal_choice_traces=882,
                  queries_required_by_formal_tasks=sum(int(r["K"]) for r in raw))
    notes.append("All formal choices match prerecommendations; each group has seven records in the frozen randomized order.")

    summary = unique_index(csv_file("summary.csv"), ("dataset", "K", "scheme"), 126, "summary")
    for (name, k, scheme), rows in groups.items():
        row = summary[name, str(k), scheme]
        for field, expected in stats([int(r["total_ns"]) / 1e6 for r in rows]).items():
            same(row[field], expected, f"summary {name}/{k}/{scheme}/{field}")
        same(row["chosen_method"], rows[0]["chosen_method"], "summary chosen method")
        same(row["trials"], 7, "summary repeats")
        same(row["distribution"], cases[name]["distribution"], "summary distribution")
        same(row["seed"], cases[name]["seed"], "summary seed")

    paired = unique_index(csv_file("paired_differences.csv"),
                          ("dataset", "K", "strategy", "control", "trial"), 504, "paired differences")
    comparisons = unique_index(csv_file("comparisons.csv"),
                               ("dataset", "K", "strategy", "control"), 72, "comparisons")
    differing_estimands = 0
    for name in cases:
        for k in protocol["K"]:
            for strategy, control in protocol["comparisons"]:
                row = comparisons[name, str(k), strategy, control]
                a, b = summary[name, str(k), strategy], summary[name, str(k), control]
                differences, percentages = [], []
                for trial in range(1, 8):
                    ta = int(index[name, k, strategy, trial]["total_ns"])
                    tb = int(index[name, k, control, trial]["total_ns"])
                    saving, pct = (tb - ta) / 1e6, 100 * (tb - ta) / tb
                    differences.append(saving)
                    percentages.append(pct)
                    pair = paired[name, str(k), strategy, control, str(trial)]
                    for field, expected in (("strategy_ms", ta / 1e6), ("control_ms", tb / 1e6),
                                             ("net_saving_ms", saving), ("net_saving_percent", pct)):
                        same(pair[field], expected, f"paired {name}/{k}/{strategy}/{control}/{trial}/{field}")
                median_diff = float(b["median_ms"]) - float(a["median_ms"])
                expected = {"strategy_median_ms": float(a["median_ms"]), "control_median_ms": float(b["median_ms"]),
                            "median_difference_saving_ms": median_diff,
                            "median_difference_saving_percent": 100 * median_diff / float(b["median_ms"]),
                            "paired_median_saving_ms": statistics.median(differences),
                            "paired_q25_saving_ms": percentile(differences, .25),
                            "paired_q75_saving_ms": percentile(differences, .75),
                            "paired_min_saving_ms": min(differences), "paired_max_saving_ms": max(differences),
                            "paired_median_saving_percent": statistics.median(percentages),
                            "worst_paired_extra_ms": -min(differences),
                            "worst_paired_extra_percent": -min(percentages),
                            "worst_loss_ms_nonnegative": max(0, -min(differences)),
                            "faster_rounds": sum(d > 0 for d in differences),
                            "same_recommendation": a["chosen_method"] == b["chosen_method"]}
                for field, value in expected.items():
                    same(row[field], value, f"comparison {name}/{k}/{strategy}/{control}/{field}")
                differing_estimands += not math.isclose(median_diff, statistics.median(differences), abs_tol=1e-9)
    counts.update(summary_rows=126, comparisons=72, paired_differences=504,
                  comparisons_with_different_median_estimands=differing_estimands)
    notes.append("Summary medians/IQR/ranges, all four signs, paired differences and worst losses recomputed from raw nanoseconds.")

    oracle_rows = unique_index(csv_file("posthoc_oracle.csv"), ("dataset", "K"), 18, "posthoc oracle")
    for name in cases:
        for k in protocol["K"]:
            best_scheme = min(FIXED, key=lambda scheme: float(summary[name, str(k), scheme]["median_ms"]))
            best = summary[name, str(k), best_scheme]
            row = oracle_rows[name, str(k)]
            same(row["best_fixed_scheme"], best_scheme, "oracle chooses fixed-method median")
            same(row["best_fixed_method"], FIXED[best_scheme], "oracle method")
            for out, source in (("best_fixed_median_ms", "median_ms"), ("q25_ms", "q25_ms"), ("q75_ms", "q75_ms")):
                same(row[out], float(best[source]), f"oracle {out}")
    regrets = unique_index(csv_file("strategy_regret.csv"), ("dataset", "K", "scheme"), 90, "strategy regret")
    for (name, k, scheme), row in regrets.items():
        current = summary[name, k, scheme]
        chosen = summary[name, k, "fixed_" + current["chosen_method"]]
        oracle = oracle_rows[name, k]
        total, fixed_total, best_total = map(float, (current["median_ms"], chosen["median_ms"], oracle["best_fixed_median_ms"]))
        for field, value in {"chosen_method": current["chosen_method"], "posthoc_best": oracle["best_fixed_method"],
                             "exact_best_method_match": current["chosen_method"] == oracle["best_fixed_method"],
                             "strategy_median_ms": total, "oracle_median_ms": best_total,
                             "extra_vs_oracle_ms": total - best_total, "chosen_fixed_penalty_ms": fixed_total - best_total,
                             "policy_minus_same_fixed_ms": total - fixed_total}.items():
            same(row[field], value, f"regret {name}/{k}/{scheme}/{field}")
    counts.update(posthoc_oracle_rows=18, strategy_regret_rows=90)

    features = unique_index(csv_file("sample_feature_errors.csv"), ("dataset",), 6, "sample feature errors")
    exact_clipped = sample_clipped = 0
    for (name,), row in features.items():
        n, numerator = int(row["N"]), int(row["r7_exact_points"])
        require(n == 10032 and 0 <= numerator <= n, "exact-f count bounds")
        exact_f = numerator / n
        sample = pre[name, 100, "sample_two"]["choice"]
        for field, value in {"exact_f": exact_f, "sample_numerator": sample["sample_numerator"],
                             "sample_f_raw": sample["f_raw"], "sample_f_used": sample["f_used"],
                             "raw_error_percentage_points": 100 * (sample["f_raw"] - exact_f),
                             "used_error_percentage_points": 100 * (sample["f_used"] - exact_f)}.items():
            same(row[field], value, f"feature error {name}/{field}")
        exact_clipped += not model["f_min"] <= exact_f <= model["f_max"]
        sample_clipped += sample["f_raw"] != sample["f_used"]
    predictions = unique_index(csv_file("fixed_prediction_errors.csv"), ("dataset", "K", "method"),
                               72, "fixed prediction errors")
    for (name, k_text, method), row in predictions.items():
        k = int(k_text)
        exact_f = float(features[name,]["exact_f"])
        actual = float(summary[name, k_text, "fixed_" + method]["median_ms"])
        predicted = model["common_geometry_ms"] + scores(model, exact_f, k)[method]
        sampled = None if k == 1 else model["common_geometry_ms"] + scores(
            model, pre[name, k, "sample_two"]["choice"]["f_raw"], k)[method]
        for field, value in {"fixed_measured_ms": actual, "exact_f_predicted_ms": predicted,
                             "exact_f_error_percent": 100 * (predicted - actual) / actual,
                             "sample_f_predicted_ms": sampled,
                             "sample_f_error_percent": None if sampled is None else 100 * (sampled - actual) / actual}.items():
            same(row[field], value, f"clipped frozen prediction {name}/{k}/{method}/{field}")
    diagnoses = unique_index(csv_file("posthoc_diagnostics.csv"), ("dataset", "K", "policy"),
                             54, "posthoc diagnostics")
    triggers = 0
    for (name, k_text, policy), row in diagnoses.items():
        k = int(k_text)
        choice = pre[name, k, policy]["choice"]
        exact_f = float(features[name,]["exact_f"])
        fixed_cost = {m: float(summary[name, k_text, "fixed_" + m]["median_ms"]) for m in METHODS}
        best = oracle_rows[name, k_text]
        new_selected = choice["selected"] in ("r6", "r8")
        if policy == "sample_four":
            triggers += new_selected
        for field, value in {"selected": choice["selected"], "base": choice["base"],
                             "raw_sample_argmin": choice["raw_argmin"], "posthoc_best": best["best_fixed_method"],
                             "exact_f": exact_f, "sample_f_raw": choice["f_raw"], "sample_f_used": choice["f_used"],
                             "new_resolution_selected": new_selected,
                             "chosen_fixed_penalty_ms": fixed_cost[choice["selected"]] - float(best["best_fixed_median_ms"])}.items():
            same(row[field], value, f"diagnostic {name}/{k}/{policy}/{field}")
        if k != 1 and policy != "k_only":
            values = scores(model, exact_f, k)
            counter = decision(values, policy == "sample_four")["selected"]
            allowed = ("direct", "r7") if policy == "sample_two" else METHODS
            raw_allowed, raw_all = min(allowed, key=values.get), min(METHODS, key=values.get)
            for field, value in {"exact_f_counterfactual_choice": counter, "exact_f_raw_allowed_min": raw_allowed,
                                 "sampling_changed_choice": choice["selected"] != counter,
                                 "sampling_fixed_cost_impact_ms": fixed_cost[choice["selected"]] - fixed_cost[counter],
                                 "model_raw_fixed_penalty_ms": fixed_cost[raw_all] - float(best["best_fixed_median_ms"]),
                                 "candidate_restriction_penalty_ms": fixed_cost[raw_allowed] - fixed_cost[raw_all],
                                 "conservative_fixed_cost_impact_ms": fixed_cost[counter] - fixed_cost[raw_allowed]}.items():
                same(row[field], value, f"clipped exact-f counterfactual {name}/{k}/{policy}/{field}")
    counts.update(frozen_prediction_rows=72, posthoc_diagnostic_rows=54,
                  exact_f_cases_clipped_by_frozen_predictor=exact_clipped,
                  sample_f_cases_clipped=sample_clipped, new_resolution_trigger_scenarios=triggers)
    notes.append("Prediction errors and exact-f counterfactuals use the frozen clipping rule; no model fitting or query is executed.")

    status = json_file(RUN / "status.json")
    for field, value in {"formal_records": 882, "prerecommendations": 54, "datasets": 6, "scenarios": 18,
                         "old_files_unchanged": True, "code_model_protocol_unchanged": True,
                         "additional_seeds_or_experiments": False, "new_resolution_triggers": triggers}.items():
        same(status[field], value, f"completion status {field}")
    require((RUN / "sha256_manifest.json").read_bytes() == original_manifest_bytes, "audit must not change original manifest")
    extra_files = sorted(p.relative_to(RUN).as_posix() for p in RUN.rglob("*") if p.is_file()
                         and p.name != "sha256_manifest.json" and p.relative_to(RUN).as_posix() not in manifest)
    delivery = RUN / "delivery_sha256_manifest.json"
    if delivery.is_file():
        counts["delivery_manifest_files"] = verify_manifest(RUN, json_file(delivery))
    return {"status": "passed", "verified_at": datetime.now().astimezone().isoformat(),
            "run": str(RUN), "audit_script_sha256": digest(Path(__file__)),
            "original_manifest_sha256": hashlib.sha256(original_manifest_bytes).hexdigest(),
            "counts": counts, "checks": notes, "files_outside_original_manifest": extra_files,
            "scope": "Read saved CSV/JSON/source bytes only; no experiment imports, point generation, geometry queries, benchmarks or refitting."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write-audit", action="store_true", help="新增审计JSON和脚本副本；不覆盖已有文件")
    args = parser.parse_args()
    targets = (RUN / "artifact_validation.json", RUN / "postprocessing" / Path(__file__).name)
    if args.write_audit and any(path.exists() for path in targets):
        parser.error("审计输出已存在；请使用默认只读模式，禁止覆盖旧审计记录")
    try:
        result = audit()
    except Exception as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1
    if args.write_audit:
        targets[1].parent.mkdir(parents=True, exist_ok=True)
        targets[1].write_bytes(Path(__file__).read_bytes())
        result["new_audit_files"] = [path.relative_to(RUN).as_posix() for path in targets]
        targets[0].write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
