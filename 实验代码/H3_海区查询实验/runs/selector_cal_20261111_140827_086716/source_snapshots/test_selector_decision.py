"""仅测试纯函数与解析阈值，不读取验证数据、不运行性能实验。"""
import unittest

from selector_decision import decide, decision_intervals, predict_costs


def model_for(lines, low=0.0, high=1.0):
    return {"f_min": low, "f_max": high,
            "methods": {method: {"prepare_ms": 0.0, "a_ms": intercept,
                                  "b_ms": slope, "mean_query_ms": intercept + slope * 0.5}
                        for method, (intercept, slope) in lines.items()}}


TOY_LINES = {"direct": (100.0, 0.0), "r6": (130.0, -80.0),
             "r7": (120.0, -40.0), "r8": (90.0, 20.0)}


class DecisionTests(unittest.TestCase):
    def test_prediction_uses_preparation_K_and_clipping(self):
        model = model_for(TOY_LINES, 0.2, 0.8)
        model["methods"]["r6"]["prepare_ms"] = 7
        self.assertAlmostEqual(predict_costs(model, -1, 3)["r6"], 7 + 3 * (130 - 80 * 0.2))
        self.assertEqual(predict_costs(model, -1, 3), predict_costs(model, 0.2, 3))
        self.assertEqual(predict_costs(model, 10, 3), predict_costs(model, 0.8, 3))

    def test_base_equality_keeps_direct_and_strict_extra_equality_does_not_switch(self):
        costs = {"direct": 102.0, "r7": 100.0, "r6": 99.96, "r8": 120.0}
        result = decide(costs, True)
        self.assertEqual(result["base"], "direct")
        self.assertEqual(result["selected"], "direct")
        self.assertEqual(result["extra_qualified"], [])
        costs["direct"] += 1e-8
        self.assertEqual(decide(costs, False)["base"], "r7")

    def test_two_candidate_control_ignores_qualified_extras_and_extra_tie_prefers_r6(self):
        costs = {"direct": 100.0, "r7": 105.0, "r6": 97.0, "r8": 97.0}
        two = decide(costs, False)
        four = decide(costs, True)
        self.assertEqual(two["selected"], "direct")
        self.assertEqual(two["extra_qualified"], ["r6", "r8"])
        self.assertEqual(four["selected"], "r6")

    def test_raw_equal_cost_priority_does_not_override_conservative_decision(self):
        costs = {"direct": 200.0, "r7": 10.0, "r6": 10.0, "r8": 10.0}
        result = decide(costs, True)
        self.assertEqual(result["raw_argmin"], "r7")
        self.assertEqual(result["selected"], "r7")
        self.assertFalse(result["extra_possible"])

    def test_all_required_roots_are_solved_analytically(self):
        result = decision_intervals(model_for(TOY_LINES), 1)
        expected = [0, 0.25, 12.4 / 40.8, 0.375, 0.4, 27.6 / 59.2, 0.5, 22.4 / 40.8, 1]
        actual = [point["f"] for point in result["boundaries"]]
        self.assertEqual(len(actual), len(expected))
        for observed, reference in zip(actual, expected):
            self.assertAlmostEqual(observed, reference, places=12)
        root = next(point for point in result["boundaries"] if abs(point["f"] - 0.4) < 1e-12)
        self.assertEqual(root["sample2"]["selected"], "direct")
        self.assertEqual(root["sample4"]["selected"], "direct")
        self.assertEqual(root["raw_argmin"], "r6")
        self.assertFalse(root["extra_possible"])
        model = model_for(TOY_LINES)
        self.assertEqual(decide(predict_costs(model, 0.39, 1), True)["selected"], "r8")
        self.assertEqual(decide(predict_costs(model, 0.41, 1), True)["selected"], "r6")

    def test_open_intervals_partition_domain_and_keep_representative_decisions(self):
        model = model_for(TOY_LINES)
        result = decision_intervals(model, 1)
        self.assertEqual(len(result["open_intervals"]), len(result["boundaries"]) - 1)
        for interval in result["open_intervals"]:
            self.assertLess(interval["left"], interval["right"])
            for fraction in (0.1, 0.5, 0.9):
                f = interval["left"] + fraction * (interval["right"] - interval["left"])
                costs = predict_costs(model, f, 1)
                self.assertEqual(decide(costs, False)["selected"], interval["sample2"]["selected"])
                self.assertEqual(decide(costs, True)["selected"], interval["sample4"]["selected"])
                self.assertEqual(decide(costs, True)["raw_argmin"], interval["raw_argmin"])

    def test_clipped_tails_are_constant_endpoint_decisions(self):
        model = model_for(TOY_LINES, 0.2, 0.8)
        result = decision_intervals(model, 1)
        for side, f in (("left", -5), ("right", 5)):
            costs = predict_costs(model, f, 1)
            self.assertEqual(result["clipped_tails"][side]["sample4"], decide(costs, True))

    def test_zero_width_range_and_parallel_lines_are_well_defined(self):
        same_lines = {method: (5.0, 2.0) for method in TOY_LINES}
        result = decision_intervals(model_for(same_lines), 1)
        self.assertEqual(len(result["open_intervals"]), 1)
        self.assertEqual(result["open_intervals"][0]["sample4"]["selected"], "direct")
        self.assertTrue(result["coincident_equations"])
        point_result = decision_intervals(model_for(same_lines, 0.3, 0.3), 0)
        self.assertEqual(len(point_result["boundaries"]), 1)
        self.assertEqual(point_result["open_intervals"], [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
