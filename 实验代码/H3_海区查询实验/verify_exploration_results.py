"""从保存的原始记录重新核对探索汇总与指纹，不进行性能测量。"""
import csv
import hashlib
import json
from pathlib import Path
import statistics

import exploration_experiment as exp

ROOT = Path(__file__).resolve().parent


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream))


def main():
    run = ROOT / (ROOT / "LATEST_EXPLORATION.txt").read_text("utf-8").strip()
    protocol = json.loads((run / "protocol.json").read_text("utf-8"))
    protocol_record = json.loads((run / "protocol_record.json").read_text("utf-8"))
    assert hashlib.sha256((run / "protocol.json").read_bytes()).hexdigest() == protocol_record["protocol_sha256"]
    assert (run / "protocol.json").read_bytes() == exp.PROTOCOL_PATH.read_bytes()
    raw = read_csv(run / "whole_batch_trials.csv")
    feature = read_csv(run / "feature_trials.csv")
    methods = read_csv(run / "whole_batch_summary.csv")
    comparisons = read_csv(run / "posthoc_comparisons.csv")
    fsummary = read_csv(run / "feature_summary.csv")
    assert len(raw) == 336 and len(feature) == 124 and len(methods) == 48 and len(comparisons) == 12
    for r in raw:
        assert int(r["whole_total_ns"]) == int(r["preparation_external_ns"]) + int(r["continuous_queries_ns"])
        assert int(r["query_calls"]) == int(r["K"])
        assert int(r["prepare_total_ns"]) == sum(int(r[k]) for k in (
            "geometry_prepare_ns", "point_index_ns", "cover_generation_ns", "cell_classification_ns", "prepare_other_ns"))
    for s in methods:
        rows = [r for r in raw if r["dataset"] == s["dataset"] and r["K"] == s["K"] and r["method"] == s["method"]]
        assert len(rows) == 7
        stats = exp.distribution_stats([int(r["whole_total_ns"]) for r in rows])
        assert all(abs(float(s[k]) - v) < 1e-12 for k, v in stats.items())
    for r in comparisons:
        candidates = [s for s in methods if s["dataset"] == r["dataset"] and s["K"] == r["K"]]
        two = min((s for s in candidates if s["method"] in ("direct", "h3_r7")), key=lambda s: float(s["median_ms"]))
        four = min(candidates, key=lambda s: float(s["median_ms"]))
        assert r["two_choice_method"] == two["method"] and r["four_choice_method"] == four["method"]
        assert abs(float(r["potential_saving_ms"]) - (float(two["median_ms"]) - float(four["median_ms"]))) < 1e-12
    for f in fsummary:
        rows = [r for r in feature if r["dataset"] == f["dataset"]]
        assert len(rows) == 31
        assert statistics.median(int(r["feature_only_ns"]) for r in rows) / 1e6 == float(f["median_ms"])
        for r in rows:
            sample_ids = json.loads(r["sample_record_ids"])
            assert len(sample_ids) == len(set(sample_ids)) == 200
            components = [int(r["sample_and_gather_ns"]), int(r["feature_other_ns"])]
            for resolution in (6, 7, 8):
                components += [int(r[f"r{resolution}_{stage}_ns"]) for stage in (
                    "encoding", "deduplication", "polygon_and_classification")]
            assert sum(components) == int(r["feature_only_ns"])
    for case in protocol["datasets"]:
        folder = run / "datasets" / case["name"]
        points = read_csv(folder / "data" / "simulated_points.csv")
        assert len(points) == len({p["record_id"] for p in points}) == 10032
        x = exp.np.array([float(p["longitude"]) for p in points])
        y = exp.np.array([float(p["latitude"]) for p in points])
        ids = exp.np.array([int(p["record_id"]) for p in points])
        region = exp.shapely.Polygon(exp.base.VERTICES)
        oracle = set(map(int, ids[exp.shapely.covers(region, exp.shapely.points(x, y))]))
        for method in exp.base.METHODS:
            selected = [int(r["record_id"]) for r in read_csv(folder / "selected_ids" / f"{method}.csv")]
            assert set(selected) == oracle and len(selected) == len(oracle)
        probes = read_csv(folder / "boundary_cases.csv")
        assert len(probes) == 32 and all(p["passed"] == "True" for p in probes)
        if case["distribution"] == "boundary_cluster":
            for label, expected in (("boundary_band_inside", 4000), ("boundary_band_outside", 4000),
                                    ("uniform_component", 2000)):
                assert sum(p["category"] == label for p in points) == expected
    manifest = json.loads((run / "sha256_manifest.json").read_text("utf-8"))
    assert all(hashlib.sha256((run / name).read_bytes()).hexdigest() == digest for name, digest in manifest.items())
    previous = json.loads((run / "previous_files_sha256.json").read_text("utf-8"))
    assert all(hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == digest for name, digest in previous.items())
    validation = {"status": "passed", "whole_batch_records": len(raw), "feature_records": len(feature),
                  "all_medians_and_variation_recomputed": True, "all_12_posthoc_comparisons_recomputed": True,
                  "saved_coordinates_rechecked_against_saved_ID_sets": True,
                  "protocol_and_output_hashes_verified": True, "previous_files_verified": len(previous)}
    exp.base.dump_json(run / "artifact_validation.json", validation)
    # 只把新产生的校验文件加入指纹清单；已有文件不改写。
    manifest["artifact_validation.json"] = hashlib.sha256((run / "artifact_validation.json").read_bytes()).hexdigest()
    exp.base.dump_json(run / "sha256_manifest.json", manifest)
    print(json.dumps(validation, ensure_ascii=False))


if __name__ == "__main__":
    main()
