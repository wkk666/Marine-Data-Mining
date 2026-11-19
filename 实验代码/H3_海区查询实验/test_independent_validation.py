"""只使用旧数据和人工时间值，绝不提前生成预留的独立验证数据。"""
import json
from unittest import TestCase,main

import run_independent_validation as v


class ValidationHarnessTests(TestCase):
    def test_locked_model_and_generation_quotas_without_generating(self):
        p=json.loads(v.PROTOCOL.read_text("utf-8"))
        self.assertEqual(v.sha(v.ROOT/p["model_source"]),p["model_sha256"])
        self.assertEqual(len(p["datasets"])*len(p["K"])*len(p["executions"])*p["repeats"],882)
        for case in p["datasets"]:
            _,cfg=v.generation_spec(case,p)
            q=cfg["boundary_cluster"]
            self.assertEqual(q["uniform_component_points"]+q["inside_band_points"]+q["outside_band_points"],10000)
            self.assertEqual(q["max_distance_to_region_boundary_degrees"],.002)

    def test_saved_choice_comparison_checks_full_feature_and_costs(self):
        model=json.loads((v.ROOT/"runs"/"selector_cal_20261111_140827_086716"/"online_model.json").read_text("utf-8"))
        from bbox_ablation import load_saved_data
        data=load_saved_data(v.ROOT/"runs"/"explore_20261110_221236_751069"/"datasets"/"uniform_A")
        raw=v.engine.RawCoordinates(data.lon,data.lat)
        a=v.engine.choose(raw,v.engine.prepare_common(),100,model,"sample_two")
        b=v.engine.choose(raw,v.engine.prepare_common(),100,model,"sample_two")
        self.assertTrue(v.choice_matches(a,b))
        altered=dict(b,sample_numerator=b["sample_numerator"]+1)
        self.assertFalse(v.choice_matches(a,altered))
        k1=v.engine.choose(raw,v.engine.prepare_common(),1,model,"sample_four")
        self.assertIsNone(k1["f_raw"]);self.assertEqual(k1["sample_rows"],[])

    def test_paired_median_is_not_confused_with_difference_of_medians(self):
        # 人工成本例子，不执行任何查询或性能实验。
        p=json.loads(v.PROTOCOL.read_text("utf-8"))
        p=dict(p,datasets=[dict(name="old_fixture",distribution="fixture",seed=0)],K=[1])
        control=[1,2,3,4,5,6,100]
        policy=[1,1,1,11,12,13,14]
        rows=[]
        for t in range(1,8):
            for s in p["executions"]:
                value=control[t-1] if s=="k_only" else policy[t-1] if s=="sample_two" else 10
                rows.append({"dataset":"old_fixture","K":1,"scheme":s,"trial":t,"total_ns":value*1000000,
                             "chosen_method":v.FIXED.get(s,"direct")})
        _,paired,comparisons,_,_=v.summarize(rows,p)
        r=next(r for r in comparisons if r["strategy"]=="sample_two" and r["control"]=="k_only")
        self.assertNotEqual(r["median_difference_saving_ms"],r["paired_median_saving_ms"])
        self.assertEqual(len(paired),28)


if __name__=="__main__":main(verbosity=2)
