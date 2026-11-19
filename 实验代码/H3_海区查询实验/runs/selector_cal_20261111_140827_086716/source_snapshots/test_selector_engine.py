import json
from unittest import TestCase, main, mock

import bbox_ablation as previous
import selector_engine as engine
from calibrate_selector import measure_fixed


class EngineTests(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = previous.ROOT
        cls.source = cls.root / "runs" / "explore_20261110_221236_751069" / "datasets"
        cls.data = previous.load_saved_data(cls.source / "boundary_A")
        cls.model = {"f_min": 0.08, "f_max": 0.54, "sample_size": 200, "sample_seed": 73001,
                     "methods": {m: {"prepare_ms": p, "a_ms": a, "b_ms": b, "mean_query_ms": a + 0.3 * b}
                                 for m, p, a, b in (("direct", .001, .2, .5), ("r6", 6, .1, .6),
                                                    ("r7", 8, .04, .7), ("r8", 19, .1, .8))}}

    def test_common_geometry_reused_in_fixed_and_selected_tasks(self):
        with mock.patch.object(engine.base, "make_region", wraps=engine.base.make_region) as create:
            common = engine.prepare_common()
            choice = engine.choose(engine.RawCoordinates(self.data.lon, self.data.lat), common, 100,
                                   self.model, "sample_four")
            state, _ = engine.prepare_specific(choice["selected"], self.data, common)
            self.assertEqual(create.call_count, 1)
            self.assertIs(state.region, common.region)
        with mock.patch.object(engine.base, "make_region", wraps=engine.base.make_region) as create:
            measure_fixed("r7", self.data, 3)
            self.assertEqual(create.call_count, 1)

    def test_feature_uses_point_weights_and_cannot_build_full_assets(self):
        raw = engine.RawCoordinates(self.data.lon, self.data.lat)
        common = engine.prepare_common()
        with mock.patch.object(engine.base, "build_point_index", side_effect=AssertionError("forbidden full index")), \
             mock.patch.object(engine.base, "enumerate_cover", side_effect=AssertionError("forbidden full cover")):
            feature = engine.r7_feature(raw, common)
        codes = previous.read_csv(self.source / "boundary_A" / "point_indexes" / "h3_r7_encoding.csv")
        classes = previous.read_csv(self.source / "boundary_A" / "coverage" / "h3_r7_cells.csv")
        border = {r["h3_cell"] for r in classes if r["class"] == "border"}
        xmin, ymin, xmax, ymax = common.bounds
        expected = sum(codes[i]["h3_cell"] in border and xmin <= self.data.lon[i] <= xmax
                       and ymin <= self.data.lat[i] <= ymax for i in feature["sample_rows"])
        self.assertEqual(feature["sample_numerator"], expected)
        self.assertEqual(feature["f_raw"], expected / 200)
        self.assertLess(feature["unique_sample_cells"], 200)
        with self.assertRaises(TypeError):
            engine.choose(self.data, common, 100, self.model, "sample_four")

    def test_controlled_two_four_features_and_costs_are_identical(self):
        raw = engine.RawCoordinates(self.data.lon, self.data.lat)
        common = engine.prepare_common()
        for k in (1, 100, 1000):
            a = engine.choose(raw, common, k, self.model, "sample_two")
            b = engine.choose(raw, common, k, self.model, "sample_four")
            self.assertEqual(a["sample_rows"], b["sample_rows"])
            self.assertEqual(a["costs"], b["costs"])
            self.assertEqual(a["base"], b["base"])
            self.assertTrue(a["selected"] == b["selected"] or b["selected"] in ("r6", "r8"))
            if k == 1:
                self.assertEqual(a["sample_size"], 0)
                self.assertIsNone(a["raw_argmin"])

    def test_unified_query_paths_equal_exact_saved_ids(self):
        for name in ("uniform_A", "boundary_A"):
            data = previous.load_saved_data(self.source / name)
            oracle = {int(r["record_id"]) for r in previous.read_csv(self.source / name / "selected_ids" / "direct.csv")}
            common = engine.prepare_common()
            for method in engine.METHODS:
                state, _ = engine.prepare_specific(method, data, common)
                ids, _ = engine.query_function(method)(state, data)
                self.assertEqual(set(map(int, ids)), oracle)
                self.assertEqual(len(ids), len(oracle))


if __name__ == "__main__":
    main(verbosity=2)
