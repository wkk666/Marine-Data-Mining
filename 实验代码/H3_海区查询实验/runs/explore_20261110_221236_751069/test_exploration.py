"""探索脚本的生成规则、成本隔离和连续执行契约测试，不选择胜出参数。"""
import json
from unittest import TestCase, main, mock

import exploration_experiment as exp


class ExplorationTests(TestCase):
    def setUp(self):
        self.protocol = json.loads(exp.PROTOCOL_PATH.read_text(encoding="utf-8"))

    def test_new_uniform_seed_preserves_original_probes(self):
        case = self.protocol["datasets"][0]
        data, _ = exp.generate_dataset(case, self.protocol)
        old_probes = exp.base.generate_data(0)
        exp.np.testing.assert_array_equal(data.lon[-32:], old_probes.lon)
        exp.np.testing.assert_array_equal(data.lat[-32:], old_probes.lat)
        self.assertEqual(len(data.ids), 10032)

    def test_boundary_generation_matches_locked_counts_and_band(self):
        data, generation = exp.generate_dataset(self.protocol["datasets"][2], self.protocol)
        region = exp.shapely.Polygon(exp.base.VERTICES)
        edge = exp.shapely.boundary(region)
        for label, expected, covered in (("boundary_band_inside", 4000, True),
                                         ("boundary_band_outside", 4000, False)):
            rows = [i for i, c in enumerate(data.category) if c == label]
            self.assertEqual(len(rows), expected)
            xy = exp.np.column_stack((data.lon[rows], data.lat[rows]))
            self.assertTrue(exp.np.all(exp.shapely.distance(edge, exp.shapely.points(xy)) <= 0.002))
            self.assertTrue(exp.np.all(exp.shapely.intersects_xy(region, xy[:, 0], xy[:, 1]) == covered))
        self.assertEqual(generation["uniform_component_points"], 2000)

    def test_feature_extraction_does_not_use_full_assets(self):
        data = exp.base.generate_data(400, 999)
        with mock.patch.object(exp.base, "build_point_index", side_effect=AssertionError("full index forbidden")), \
             mock.patch.object(exp.base, "enumerate_cover", side_effect=AssertionError("full cover forbidden")):
            rows, metrics = exp.feature_extraction(data, exp.base.make_region(), 200, 111)
        self.assertEqual(len(set(map(int, rows))), 200)
        self.assertGreater(metrics["feature_only_ns"], 0)
        for r in (6, 7, 8):
            self.assertEqual(metrics[f"r{r}_unique_cells"],
                             sum(metrics[f"r{r}_sample_cells_{x}"] for x in ("internal", "border", "outside")))

    def test_batch_executes_k_queries_and_new_results_are_constructed(self):
        data = exp.base.generate_data(100, 1001)
        for method in exp.base.METHODS:
            with self.subTest(method=method):
                with mock.patch.object(exp.base, "query", wraps=exp.base.query) as query:
                    ids, metrics = exp.measure_batch(method, data, 5)
                    self.assertEqual(query.call_count, 5)
                self.assertEqual(metrics["whole_total_ns"], metrics["preparation_external_ns"]
                                 + metrics["continuous_queries_ns"])
                state, _ = exp.base.prepare_method(method, data)
                one, _ = exp.base.query(state, data)
                two, _ = exp.base.query(state, data)
                self.assertFalse(exp.np.shares_memory(one, two))
                self.assertEqual(set(map(int, ids)), set(map(int, one)))


if __name__ == "__main__":
    main(verbosity=2)
