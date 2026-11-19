import json
from unittest import TestCase, main, mock

import bbox_ablation as experiment


class BboxTests(TestCase):
    def setUp(self):
        self.protocol = json.loads(experiment.PROTOCOL_PATH.read_text("utf-8"))
        self.source = experiment.ROOT / self.protocol["source_run"]

    def test_ids_and_boundary_semantics_on_all_saved_datasets(self):
        b = experiment.base
        for name in self.protocol["datasets"]:
            data = experiment.load_saved_data(self.source / "datasets" / name)
            region = b.shapely.Polygon(b.VERTICES)
            oracle = set(map(int, data.ids[b.shapely.covers(region, b.shapely.points(data.lon, data.lat))]))
            for resolution in (6, 7, 8):
                with self.subTest(dataset=name, resolution=resolution):
                    state, _ = b.prepare_method(f"h3_r{resolution}", data)
                    old_ids, old_metrics = b.query(state, data)
                    for profile in (False, True):
                        ids, metrics = experiment.query_h3_bbox(state, data, profile=profile)
                        self.assertEqual(set(map(int, ids)), oracle)
                        self.assertEqual(len(ids), len(oracle))
                        self.assertEqual(metrics["internal_accepted_points"], old_metrics["internal_accepted_points"])
                    if name == "boundary_A" and resolution == 7:
                        self.assertEqual(old_metrics["exact_points"], 8198)
                        self.assertEqual(metrics["bbox_removed_points"], 2965)
                        self.assertEqual(metrics["exact_points"], 5233)

    def test_batch_calls_variant_k_times_and_returns_new_arrays(self):
        data = experiment.load_saved_data(self.source / "datasets" / "boundary_A")
        for resolution in (6, 7, 8):
            method = f"h3_r{resolution}_bbox"
            with mock.patch.object(experiment, "query_h3_bbox", wraps=experiment.query_h3_bbox) as query:
                ids, metrics = experiment.measure_batch(method, data, 3)
                self.assertEqual(query.call_count, 3)
            self.assertEqual(metrics["whole_total_ns"], metrics["preparation_external_ns"] + metrics["continuous_queries_ns"])
            state, _ = experiment.base.prepare_method(experiment.original_method(method), data)
            first, _ = experiment.query_h3_bbox(state, data)
            second, _ = experiment.query_h3_bbox(state, data)
            self.assertFalse(experiment.np.shares_memory(first, second))
            self.assertEqual(set(map(int, ids)), set(map(int, first)))


if __name__ == "__main__":
    main(verbosity=2)
