"""核对边界规则、非连续记录 ID 和三档查询；不测试性能结论。"""
import unittest

import run_experiment as ex

np, shapely = ex.np, ex.shapely


class CorrectnessTests(unittest.TestCase):
    def test_boundary_and_nearby_points_have_independent_expected_answers(self):
        data = ex.generate_data(1)
        region = shapely.Polygon(ex.VERTICES)
        answer = shapely.covers(region, shapely.points(data.lon, data.lat))
        touches = shapely.intersects_xy(shapely.boundary(region), data.lon, data.lat)
        probes = 0
        for row, expected in enumerate(data.expected_probe_membership):
            if expected is not None:
                probes += 1
                self.assertEqual(bool(answer[row]), expected, data.category[row])
                self.assertEqual(bool(touches[row]), data.category[row].startswith("on_"),
                                 data.category[row])
        self.assertEqual(probes, 32)

    def test_methods_return_record_ids_with_shuffled_nonconsecutive_ids(self):
        data = ex.generate_data(600)
        rng = np.random.default_rng(42)
        data.ids = rng.permutation(np.arange(len(data.ids), dtype=np.int64) * 17 + 100003)
        oracle = set(map(int, data.ids[shapely.covers(shapely.Polygon(ex.VERTICES),
                                                   shapely.points(data.lon, data.lat))]))
        for method in ex.METHODS:
            with self.subTest(method=method):
                state, _ = ex.prepare_method(method, data)
                ids, metrics = ex.query(state, data)
                self.assertEqual(set(map(int, ids)), oracle)
                self.assertEqual(len(ids), len(set(map(int, ids))))
                self.assertEqual(metrics["query_total_ns"], metrics["lookup_and_gather_ns"]
                                 + metrics["exact_predicate_ns"] + metrics["assemble_result_ns"])
                if method != "direct":
                    self.assertEqual(set(state.internal) | set(state.border), set(state.cover))
                    self.assertFalse(set(state.internal) & set(state.border))
                    self.assertTrue(ex.check_coverage(state)["passed"])

    def test_query_uses_current_coordinates_in_border_cells_instead_of_caching_answers(self):
        data = ex.generate_data(1)
        for method in ex.METHODS:
            state, _ = ex.prepare_method(method, data)
            # 同一海区边中点的内、外近邻通常同格；在原索引不变时切换经纬度。
            for edge in range(8):
                inside_row, outside_row = 2 + edge * 3, 3 + edge * 3
                if state.encoded is not None and state.encoded[inside_row] != state.encoded[outside_row]:
                    continue
                before, _ = ex.query(state, data)
                original_lon, original_lat = data.lon[inside_row], data.lat[inside_row]
                data.lon[inside_row], data.lat[inside_row] = data.lon[outside_row], data.lat[outside_row]
                try:
                    after, _ = ex.query(state, data)
                    self.assertIn(int(data.ids[inside_row]), before)
                    self.assertNotIn(int(data.ids[inside_row]), after)
                finally:
                    data.lon[inside_row], data.lat[inside_row] = original_lon, original_lat


if __name__ == "__main__":
    unittest.main(verbosity=2)
