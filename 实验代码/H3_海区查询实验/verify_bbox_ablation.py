"""复核保存的输入、ID、候选计数、原始计时汇总与旧文件指纹，不跑性能实验。"""
import hashlib
import json
from pathlib import Path
import statistics

import bbox_ablation as experiment

ROOT = Path(__file__).resolve().parent


def main():
    run = ROOT / (ROOT / "LATEST_BBOX_ABLATION.txt").read_text("utf-8").strip()
    read = experiment.read_csv
    protocol = json.loads((run / "protocol.json").read_text("utf-8"))
    assert (run / "protocol.json").read_bytes() == experiment.PROTOCOL_PATH.read_bytes()
    raw = read(run / "whole_batch_trials.csv")
    profiles = read(run / "profile_trials.csv")
    totals = read(run / "whole_batch_summary.csv")
    phases = read(run / "profile_summary.csv")
    pairs = read(run / "pair_comparisons.csv")
    ranks = read(run / "rankings.csv")
    assert len(raw) == 196 and len(profiles) == 868 and len(totals) == 28 and len(pairs) == 12
    for r in raw:
        assert int(r["K"]) == int(r["query_calls"]) == 1000
        assert int(r["whole_total_ns"]) == int(r["preparation_external_ns"]) + int(r["continuous_queries_ns"])
    for r in profiles:
        assert int(r["query_total_ns"]) == sum(int(r[k]) for k in (
            "lookup_and_gather_ns", "exact_predicate_ns", "assemble_result_ns"))
        assert int(r["bbox_filter_ns"]) == int(r["bbox_mask_ns"]) + int(r["bbox_array_compaction_ns"])
    for s in totals:
        rows = [r for r in raw if r["dataset"] == s["dataset"] and r["method"] == s["method"]]
        assert len(rows) == 7
        stats = experiment.distribution_stats([int(r["whole_total_ns"]) for r in rows])
        assert all(abs(float(s[k]) - v) < 1e-9 for k, v in stats.items())
    for s in phases:
        rows = [r for r in profiles if r["dataset"] == s["dataset"] and r["method"] == s["method"]]
        assert len(rows) == 31
        for key in ("query_total", "exact_predicate", "bbox_filter", "bbox_mask", "bbox_array_compaction"):
            assert abs(statistics.median(int(r[key + "_ns"]) for r in rows) / 1000 - float(s[key + "_median_us"])) < 1e-9
    for name in protocol["datasets"]:
        points = read(run / "input_data" / name / "simulated_points.csv")
        x = experiment.np.array([float(p["longitude"]) for p in points])
        y = experiment.np.array([float(p["latitude"]) for p in points])
        ids = experiment.np.array([int(p["record_id"]) for p in points])
        region = experiment.shapely.Polygon(experiment.base.VERTICES)
        oracle = set(map(int, ids[experiment.shapely.covers(region, experiment.shapely.points(x, y))]))
        assert len(ids) == len(set(map(int, ids))) == 10032
        for method in protocol["methods"]:
            selected = [int(row["record_id"]) for row in read(run / "datasets" / name / "selected_ids" / f"{method}.csv")]
            assert len(selected) == len(set(selected)) and set(selected) == oracle
        for resolution in (6, 7, 8):
            method = f"h3_r{resolution}"
            cells = read(run / "datasets" / name / "coverage" / f"{method}_cells.csv")
            border = {c["h3_cell"] for c in cells if c["class"] == "border"}
            encodings = read(run / "datasets" / name / "point_indexes" / f"{method}_encoding.csv")
            rows = experiment.np.array([i for i, p in enumerate(encodings) if p["h3_cell"] in border])
            xmin, ymin, xmax, ymax = region.bounds
            keep = (x[rows] >= xmin) & (x[rows] <= xmax) & (y[rows] >= ymin) & (y[rows] <= ymax)
            pair = next(p for p in pairs if p["dataset"] == name and int(p["resolution"]) == resolution)
            assert (int(pair["border_candidates"]), int(pair["bbox_removed"]), int(pair["new_exact_points"])) == (
                len(rows), int((~keep).sum()), int(keep.sum()))
            old = next(s for s in totals if s["dataset"] == name and s["method"] == method)
            new = next(s for s in totals if s["dataset"] == name and s["method"] == method + "_bbox")
            assert abs(float(pair["total_saved_ms"]) - (float(old["median_ms"]) - float(new["median_ms"]))) < 1e-9
        rank = next(r for r in ranks if r["dataset"] == name)
        order = sorted([s for s in totals if s["dataset"] == name], key=lambda s: float(s["median_ms"]))
        assert rank["all_seven_ranking"] == " < ".join(s["method"] for s in order)
    inputs = json.loads((run / "input_manifest.json").read_text("utf-8"))
    for name, info in inputs.items():
        assert (run / name).read_bytes() == (ROOT / info["source"]).read_bytes()
        assert hashlib.sha256((run / name).read_bytes()).hexdigest() == info["sha256"]
    previous = json.loads((run / "previous_files_sha256.json").read_text("utf-8"))
    assert all(hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == digest for name, digest in previous.items())
    manifest = json.loads((run / "sha256_manifest.json").read_text("utf-8"))
    assert all(hashlib.sha256((run / name).read_bytes()).hexdigest() == digest for name, digest in manifest.items())
    validation = {"status": "passed", "raw_whole_trials": 196, "raw_profile_trials": 868,
                  "all_28_method_medians_and_variation_recomputed": True, "all_12_candidate_counts_recomputed": True,
                  "all_28_saved_ID_sets_equal_to_exact_oracle": True, "input_bytes_unchanged": True,
                  "previous_files_unchanged": len(previous)}
    experiment.base.dump_json(run / "artifact_validation.json", validation)
    manifest["artifact_validation.json"] = hashlib.sha256((run / "artifact_validation.json").read_bytes()).hexdigest()
    experiment.base.dump_json(run / "sha256_manifest.json", manifest)
    print(json.dumps(validation, ensure_ascii=False))


if __name__ == "__main__":
    main()
