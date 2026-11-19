"""冻结模型的六组独立验证：先保存全部推荐，再预检和端到端实测。"""
from datetime import datetime
import hashlib
import json
from pathlib import Path
import random
import statistics
import time
import gc

import selector_engine as engine
from selector_decision import predict_costs, decide
from exploration_experiment import generate_dataset, save_dataset, distribution_stats

base, np = engine.base, engine.np
ROOT = Path(__file__).resolve().parent
PROTOCOL = ROOT / "independent_validation_protocol.json"
POLICIES = ("k_only", "sample_two", "sample_four")
FIXED = {"fixed_direct":"direct", "fixed_r7":"r7", "fixed_r6":"r6", "fixed_r8":"r8"}
FREEZE_FILES = ("run_independent_validation.py", "test_independent_validation.py", "selector_engine.py",
                "selector_decision.py", "run_experiment.py", "bbox_ablation.py", "exploration_experiment.py", "requirements.txt")


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def event(run, name, **details):
    record = {"event":name,"timestamp":datetime.now().astimezone().isoformat(),
              "monotonic_ns":time.perf_counter_ns(), **details}
    with (run / "events.jsonl").open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(record,ensure_ascii=False)+"\n")


def protected_files():
    paths = [p for p in ROOT.iterdir() if p.is_file()]
    paths += [p for p in (ROOT / "runs").rglob("*") if p.is_file()]
    return {p.relative_to(ROOT).as_posix():sha(p) for p in paths}


def generation_spec(case, protocol):
    cfg = {"random_points_per_dataset":10000, "boundary_cluster":{
        "uniform_component_points":case["uniform_points"], "inside_band_points":case["inside_band"],
        "outside_band_points":case["outside_band"], "max_distance_to_region_boundary_degrees":0.002,
        "proposal_batch_size":32768}}
    generator_case = {"name":case["name"], "seed":case["seed"],
                      "distribution":"uniform" if case["distribution"]=="uniform" else "boundary_cluster"}
    assert sum((case["uniform_points"],case["inside_band"],case["outside_band"]))==10000
    return generator_case,cfg


def save_prerecommendations(run, datasets, model, protocol):
    records, lookup = [], {}
    for name,data in datasets.items():
        raw=engine.RawCoordinates(data.lon,data.lat)
        for k in protocol["K"]:
            for policy in POLICIES:
                # 不传入数据标签、full index、准确f或测得成本；每次自己准备G。
                common=engine.prepare_common()
                choice=engine.choose(raw,common,k,model,policy)
                rows=choice["sample_rows"]
                record={"dataset":name,"K":k,"policy":policy,"sample_seed":model["sample_seed"],
                        "sample_record_ids":data.ids[rows].tolist() if rows else [],"choice":choice}
                records.append(record)
                lookup[name,k,policy]=choice
    assert len(records)==protocol["prerecommendation_count"]==54
    path=run/"prerecommendations.jsonl"
    with path.open("w",encoding="utf-8") as stream:
        for record in records:
            stream.write(json.dumps(record,ensure_ascii=False)+"\n")
    base.write_csv(run/"prerecommendation_summary.csv",[{
        "dataset":r["dataset"],"K":r["K"],"policy":r["policy"],"selected":r["choice"]["selected"],
        "base":r["choice"]["base"],"raw_argmin":r["choice"]["raw_argmin"],
        "sample_numerator":r["choice"].get("sample_numerator"),"sample_size":r["choice"].get("sample_size"),
        "f_raw":r["choice"]["f_raw"],"f_used":r["choice"]["f_used"],
        "extra_qualified":json.dumps(r["choice"].get("extra_qualified",[]))} for r in records])
    digest=sha(path)
    base.dump_json(run/"prerecommendations_lock.json",{"count":54,"sha256":digest,
        "all_saved_before_full_correctness_warmup_or_timing":True})
    event(run,"all_recommendations_locked",count=54,sha256=digest)
    return lookup


def choice_matches(choice, saved):
    keys=("selected","base","raw_argmin","f_raw","f_used","sample_rows","sample_numerator",
          "sample_size","costs","extra_qualified")
    return all(choice.get(key)==saved.get(key) for key in keys)


def preflight(run,datasets,model,protocol,choices):
    event(run,"correctness_preflight_started")
    checks, oracles = {}, {}
    for name,data in datasets.items():
        geom=base.shapely.Polygon(base.VERTICES)
        expected=set(map(int,data.ids[base.shapely.covers(geom,base.shapely.points(data.lon,data.lat))]))
        entries={}
        for method in engine.METHODS:
            common=engine.prepare_common()
            state,_=engine.prepare_specific(method,data,common)
            ids,_=engine.query_function(method)(state,data)
            actual=set(map(int,ids))
            cover=None if method=="direct" else base.check_coverage(state)
            entry={"selected_points":len(ids),"false_positive":len(actual-expected),
                   "false_negative":len(expected-actual),"duplicates":len(ids)-len(actual),
                   "boundary_probes_passed":all((int(data.ids[i]) in actual)==truth
                     for i,truth in enumerate(data.expected_probe_membership) if truth is not None),"coverage":cover}
            entries[method]=entry
            base.write_csv(run/"selected_ids"/name/f"{method}.csv",
                           [{"record_id":i} for i in sorted(actual)],["record_id"])
            if actual!=expected or entry["duplicates"] or not entry["boundary_probes_passed"] or (cover and not cover["passed"]):
                base.dump_json(run/"preflight_failure.json",{"dataset":name,"method":method,"check":entry,
                               "extra":sorted(actual-expected),"missing":sorted(expected-actual)})
                raise RuntimeError("验证预检失败，保存诊断并停止")
        for k in protocol["K"]:
            for policy in POLICIES:
                ids,choice=engine.execute_task(data,1 if k==1 else k,model=model,strategy=policy)
                actual=set(map(int,ids))
                assert choice_matches(choice,choices[name,k,policy])
                assert actual==expected and len(ids)==len(expected)
                entries[f"{policy}_K{k}"]={"selected":choice["selected"],"false_positive":0,
                    "false_negative":0,"duplicates":0,"boundary_probes_passed":True,"prerecommendation_match":True}
        checks[name],oracles[name]=entries,expected
        print(f"{name}: 固定方法/选择策略ID、32边界点及覆盖通过，返回{len(expected)}条。",flush=True)
    base.dump_json(run/"correctness.json",checks)
    event(run,"correctness_preflight_completed")
    return checks,oracles


def execute(scheme,data,k,model):
    if scheme in FIXED:
        return engine.execute_task(data,k,fixed_method=FIXED[scheme])
    return engine.execute_task(data,k,model=model,strategy=scheme)


def benchmark(run,datasets,model,protocol,choices,oracles):
    cases=[(case["name"],k,scheme) for case in protocol["datasets"] for k in protocol["K"] for scheme in protocol["executions"]]
    # 统一预热在全部推荐和正确性检查之后；没有可转移到正式任务的准备资产。
    for name in datasets:
        for scheme in protocol["executions"]:
            execute(scheme,datasets[name],100,model)
    event(run,"uniform_warmup_completed")
    rng=random.Random(protocol["order_seed"])
    records,traces=[],[]
    enabled=gc.isenabled()
    gc.disable()
    event(run,"formal_timing_started")
    try:
        for trial in range(1,protocol["repeats"]+1):
            order=cases.copy();rng.shuffle(order)
            for position,(name,k,scheme) in enumerate(order,1):
                start=time.perf_counter_ns()
                ids,choice=execute(scheme,datasets[name],k,model)
                end=time.perf_counter_ns()
                actual=set(map(int,ids))
                match=scheme in FIXED or choice_matches(choice,choices[name,k,scheme])
                if actual!=oracles[name] or len(ids)!=len(actual) or not match:
                    base.dump_json(run/"formal_failure.json",{"trial":trial,"dataset":name,"K":k,"scheme":scheme,
                        "extra":sorted(actual-oracles[name]),"missing":sorted(oracles[name]-actual),
                        "duplicates":len(ids)-len(actual),"recommendation_match":match,"choice":choice})
                    if records: base.write_csv(run/"formal_trials_before_failure.csv",records)
                    raise RuntimeError("正式任务正确性或预先推荐一致性失败；停止")
                records.append({"trial":trial,"order_in_round":position,"dataset":name,"K":k,"scheme":scheme,
                    "total_ns":end-start,"chosen_method":choice["selected"],"selected_records":len(ids),
                    "false_positive":0,"false_negative":0,"duplicates":0,"prerecommendation_match":match})
                traces.append({"trial":trial,"dataset":name,"K":k,"scheme":scheme,"choice":choice})
            print(f"独立验证正式计时：{trial}/{protocol['repeats']}轮（每轮126任务）",flush=True)
            event(run,"formal_round_completed",trial=trial,records=len(records))
    finally:
        if enabled: gc.enable()
    assert len(records)==882
    event(run,"formal_timing_completed",records=len(records))
    return records,traces


def summarize(records,protocol):
    summary,paired,comparisons,oracles,regrets=[],[],[],[],[]
    index={(r["dataset"],r["K"],r["scheme"],r["trial"]):r for r in records}
    for case in protocol["datasets"]:
        name=case["name"]
        for k in protocol["K"]:
            entries={}
            for scheme in protocol["executions"]:
                rows=[index[name,k,scheme,t] for t in range(1,8)]
                entry={"dataset":name,"distribution":case["distribution"],"seed":case["seed"],"K":k,
                       "scheme":scheme,"trials":len(rows),"chosen_method":rows[0]["chosen_method"],
                       **distribution_stats([r["total_ns"] for r in rows])}
                summary.append(entry);entries[scheme]=entry
            best=min((entries[s] for s in FIXED),key=lambda r:r["median_ms"])
            oracle={"dataset":name,"K":k,"best_fixed_scheme":best["scheme"],
                    "best_fixed_method":FIXED[best["scheme"]],"best_fixed_median_ms":best["median_ms"],
                    "q25_ms":best["q25_ms"],"q75_ms":best["q75_ms"]}
            oracles.append(oracle)
            for strategy,control in protocol["comparisons"]:
                a,b=entries[strategy],entries[control]
                diffs=[];percent=[]
                for trial in range(1,8):
                    ta=index[name,k,strategy,trial]["total_ns"];tb=index[name,k,control,trial]["total_ns"]
                    delta=(tb-ta)/1e6
                    diffs.append(delta);percent.append(100*(tb-ta)/tb)
                    paired.append({"dataset":name,"K":k,"trial":trial,"strategy":strategy,"control":control,
                                   "strategy_ms":ta/1e6,"control_ms":tb/1e6,"net_saving_ms":delta,
                                   "net_saving_percent":percent[-1]})
                med_diff=b["median_ms"]-a["median_ms"]
                stats=distribution_stats([d*1e6 for d in diffs])
                comparisons.append({"dataset":name,"K":k,"strategy":strategy,"control":control,
                    "strategy_median_ms":a["median_ms"],"control_median_ms":b["median_ms"],
                    "median_difference_saving_ms":med_diff,"median_difference_saving_percent":100*med_diff/b["median_ms"],
                    "paired_median_saving_ms":stats["median_ms"],"paired_q25_saving_ms":stats["q25_ms"],
                    "paired_q75_saving_ms":stats["q75_ms"],"paired_min_saving_ms":stats["min_ms"],
                    "paired_max_saving_ms":stats["max_ms"],"paired_median_saving_percent":statistics.median(percent),
                    "worst_paired_extra_ms":-min(diffs),"worst_paired_extra_percent":-min(percent),
                    "worst_loss_ms_nonnegative":max(0,-min(diffs)),
                    "faster_rounds":sum(d>0 for d in diffs),
                    "same_recommendation":a["chosen_method"]==b["chosen_method"]})
            for scheme in ("fixed_direct","fixed_r7",*POLICIES):
                row=entries[scheme];chosen_scheme="fixed_"+row["chosen_method"]
                fixed_chosen=entries[chosen_scheme]
                regrets.append({"dataset":name,"K":k,"scheme":scheme,"chosen_method":row["chosen_method"],
                    "posthoc_best":oracle["best_fixed_method"],"exact_best_method_match":row["chosen_method"]==oracle["best_fixed_method"],
                    "strategy_median_ms":row["median_ms"],"oracle_median_ms":oracle["best_fixed_median_ms"],
                    "extra_vs_oracle_ms":row["median_ms"]-oracle["best_fixed_median_ms"],
                    "chosen_fixed_penalty_ms":fixed_chosen["median_ms"]-oracle["best_fixed_median_ms"],
                    "policy_minus_same_fixed_ms":row["median_ms"]-fixed_chosen["median_ms"],
                    "gap_note":"policy-minus-same-fixed includes selection and execution context/noise; not isolated S measurement"})
    return summary,paired,comparisons,oracles,regrets


def posthoc_diagnostics(run,datasets,model,protocol,choices,summary,oracles):
    event(run,"posthoc_diagnostics_started")
    lookup={(r["dataset"],r["K"],r["scheme"]):r for r in summary}
    best={(r["dataset"],r["K"]):r for r in oracles}
    diagnostics,feature_errors,prediction_errors=[],[],[]
    for name,data in datasets.items():
        common=engine.prepare_common();state,_=engine.prepare_specific("r7",data,common)
        _,metrics=engine.query_function("r7")(state,data)
        exact_f=metrics["exact_points"]/len(data.ids)
        sample=choices[name,100,"sample_two"]
        feature_errors.append({"dataset":name,"N":len(data.ids),"r7_exact_points":metrics["exact_points"],
            "exact_f":exact_f,"sample_numerator":sample["sample_numerator"],"sample_f_raw":sample["f_raw"],
            "sample_f_used":sample["f_used"],"raw_error_percentage_points":100*(sample["f_raw"]-exact_f),
            "used_error_percentage_points":100*(sample["f_used"]-exact_f)})
        for k in protocol["K"]:
            fixed_cost={m:lookup[name,k,"fixed_"+m]["median_ms"] for m in engine.METHODS}
            true_scores=predict_costs(model,exact_f,k)
            for method in engine.METHODS:
                actual=fixed_cost[method]
                predicted_exact=model["common_geometry_ms"]+true_scores[method]
                sampled=None if k==1 else model["common_geometry_ms"]+predict_costs(model,sample["f_raw"],k)[method]
                prediction_errors.append({"dataset":name,"K":k,"method":method,"fixed_measured_ms":actual,
                    "exact_f_predicted_ms":predicted_exact,"exact_f_error_percent":100*(predicted_exact-actual)/actual,
                    "sample_f_predicted_ms":sampled,"sample_f_error_percent":None if sampled is None else 100*(sampled-actual)/actual,
                    "note":"posthoc diagnostic only; selection cost excluded; no refitting"})
            for policy in POLICIES:
                choice=choices[name,k,policy];selected=choice["selected"]
                entry={"dataset":name,"K":k,"policy":policy,"selected":selected,"base":choice["base"],
                    "raw_sample_argmin":choice["raw_argmin"],"posthoc_best":best[name,k]["best_fixed_method"],
                    "exact_f":exact_f,"sample_f_raw":choice["f_raw"],"sample_f_used":choice["f_used"],
                    "new_resolution_selected":selected in ("r6","r8"),
                    "chosen_fixed_penalty_ms":fixed_cost[selected]-best[name,k]["best_fixed_median_ms"]}
                if k==1 or policy=="k_only":
                    entry.update(exact_f_counterfactual_choice=None,exact_f_raw_allowed_min=None,
                        sampling_changed_choice=None,sampling_fixed_cost_impact_ms=None,
                        model_raw_fixed_penalty_ms=None,candidate_restriction_penalty_ms=None,
                        conservative_fixed_cost_impact_ms=None)
                else:
                    counter=decide(true_scores,allow_extra=policy=="sample_four")
                    allowed=("direct","r7") if policy=="sample_two" else tuple(engine.METHODS)
                    raw_exact=min(allowed,key=lambda m:true_scores[m])
                    allowed_best=min(allowed,key=lambda m:fixed_cost[m])
                    cf=counter["selected"]
                    entry.update(exact_f_counterfactual_choice=cf,exact_f_raw_allowed_min=raw_exact,
                        sampling_changed_choice=selected!=cf,
                        sampling_fixed_cost_impact_ms=fixed_cost[selected]-fixed_cost[cf],
                        model_raw_fixed_penalty_ms=fixed_cost[raw_exact]-fixed_cost[allowed_best],
                        candidate_restriction_penalty_ms=fixed_cost[allowed_best]-best[name,k]["best_fixed_median_ms"],
                        conservative_fixed_cost_impact_ms=fixed_cost[cf]-fixed_cost[raw_exact])
                entry["diagnosis_note"]="overlapping descriptive counterfactuals, not exclusive causal attribution; no feedback"
                diagnostics.append(entry)
    return diagnostics,feature_errors,prediction_errors


def write_results(run,summary,comparisons,oracles,regrets,diagnostics,features,protocol):
    lines=["# 冻结模型的六组独立验证", "",
        "全部为模拟数据，独立生成6组；K和7次重复不增加独立数据集数量。模型、200点、种子73001及2%规则未调整。", "",
        "54条推荐在任何全量正确性预检、预热和性能计时之前全部保存并锁定。正式任务重新执行并支付选择工作。", "",
        "882条正式记录由同一外层计时器测量：共同几何、选择（如有）、剩余准备和K次正常返回；无新增分项诊断探针、无最终答案缓存。", "",
        "## 18场景关键比较", "",
        "净节省 = 对照总时间中位数 − 策略总时间中位数；正数为策略更快，负数为更慢，单位毫秒。", "",
        "| 数据 | K | 样本二选一相对K-only | 四选一相对二选一 | 四选一相对直接 | 四选一相对固定r7 |",
        "| --- | ---: | ---: | ---: | ---: | ---: |"]
    comp={(r["dataset"],r["K"],r["strategy"],r["control"]):r for r in comparisons}
    for case in protocol["datasets"]:
        for k in protocol["K"]:
            vals=[comp[case["name"],k,a,b]["median_difference_saving_ms"] for a,b in protocol["comparisons"]]
            lines.append(f"| {case['name']} | {k} | "+" | ".join(f"{v:+.4f}" for v in vals)+" |")
    lines += ["", "所有净节省百分比、同轮配对差中位数、25–75百分位及最坏损失见comparisons.csv和paired_differences.csv。中位数之差与差值中位数是两种统计量，未混写。", "",
        "## 抽样比例与推荐", "",
        "| 数据 | 事后准确f | 原样本f | 裁剪后f | K100二/四选一 | K1000二/四选一 |", "| --- | ---: | ---: | ---: | --- | --- |"]
    diag={(r["dataset"],r["K"],r["policy"]):r for r in diagnostics}
    for r in features:
        n=r["dataset"]
        lines.append(f"| {n} | {r['exact_f']:.5f} | {r['sample_f_raw']:.5f} | {r['sample_f_used']:.5f} | "
            f"{diag[n,100,'sample_two']['selected']}/{diag[n,100,'sample_four']['selected']} | "
            f"{diag[n,1000,'sample_two']['selected']}/{diag[n,1000,'sample_four']['selected']} |")
    triggers=[r for r in diagnostics if r["policy"]=="sample_four" and r["new_resolution_selected"]]
    lines += ["", f"样本四选一新增档触发场景数：{len(triggers)}。"]
    if not triggers: lines += ["**本验证集未触发新增档位，未验证其被选中后的收益。**"]
    lines += ["若二选一、四选一选择同一方法，细小计时差不归因于新增分辨率。2%不是实测显著性门槛。", "",
        "## 全部方案的完整耗时", "",
        "单位毫秒。括号为第25–75百分位，最小/最大及原始7次值另见CSV。", "",
        "| 数据 | K | 方案 | 实际方法 | 中位数 | 第25–75百分位 |", "| --- | ---: | --- | --- | ---: | --- |"]
    for r in summary:
        lines.append(f"| {r['dataset']} | {r['K']} | {r['scheme']} | {r['chosen_method']} | {r['median_ms']:.5f} | {r['q25_ms']:.5f}–{r['q75_ms']:.5f} |")
    lines += ["", "## 解释边界与文件", "",
        "事后最优取四固定方法总耗时中位数的最小值，仅作参考；不是原始记录中的偶然最小值，也不是运行时可免费获取的方法。",
        "抽样、模型和保守规则的诊断标签可以重叠。策略与其选中固定方法的时间差含选择与执行上下文/波动，不是单独S的精确测量。",
        "事后准确f只在推荐锁定与正式计时结束后计算；诊断结果不修改已保存推荐、模型或参数。",
        "独立生成的数据仍然只有6组，种子相同的7次执行只是计时波动，不是7次独立抽样。",
        "文件：formal_trials.csv、formal_choice_traces.jsonl、prerecommendations.jsonl、correctness.json、comparisons.csv、paired_differences.csv、posthoc_oracle.csv、strategy_regret.csv、posthoc_diagnostics.csv、fixed_prediction_errors.csv。",
        "模型与协议及代码在snapshots/，指纹与事件顺序用于复现审计。未追加数据、未重新校准，未生成完整报告或发布GitHub。", ""]
    (run/"独立验证结果.md").write_text("\n".join(lines),encoding="utf-8")


def main():
    protocol_bytes=PROTOCOL.read_bytes();protocol=json.loads(protocol_bytes)
    protected=protected_files()
    source=ROOT/protocol["model_source"]
    model_bytes=source.read_bytes()
    assert hashlib.sha256(model_bytes).hexdigest()==protocol["model_sha256"]
    provenance={}
    for filename in ("selector_engine.py","selector_decision.py","run_experiment.py","bbox_ablation.py"):
        archived=source.parent/"source_snapshots"/filename
        assert (ROOT/filename).read_bytes()==archived.read_bytes(),f"frozen calibration code changed: {filename}"
        provenance[filename]={"source":archived.relative_to(ROOT).as_posix(),"sha256":sha(archived)}
    archived_generator=ROOT/"runs"/"explore_20261110_221236_751069"/"exploration_experiment.py"
    assert (ROOT/"exploration_experiment.py").read_bytes()==archived_generator.read_bytes()
    provenance["exploration_experiment.py"]={"source":archived_generator.relative_to(ROOT).as_posix(),"sha256":sha(archived_generator)}
    model=json.loads(model_bytes)
    assert model["sample_size"]==200 and model["sample_seed"]==73001
    assert len(protocol["datasets"])==6 and protocol["K"]==[1,100,1000] and protocol["repeats"]==7
    run=ROOT/"runs"/datetime.now().astimezone().strftime("validation_%Y%m%d_%H%M%S_%f")
    run.mkdir(parents=True)
    (run/"protocol.json").write_bytes(protocol_bytes)
    (run/"online_model.json").write_bytes(model_bytes)
    base.dump_json(run/"previous_files_sha256.json",protected)
    hashes={}
    for filename in FREEZE_FILES:
        p=run/"snapshots"/filename;p.parent.mkdir(exist_ok=True)
        p.write_bytes((ROOT/filename).read_bytes());hashes[filename]=sha(p)
    frozen={"frozen_before_generation":datetime.now().astimezone().isoformat(),
            "code_sha256":hashes,"model_sha256":sha(run/"online_model.json"),"protocol_sha256":sha(run/"protocol.json")}
    base.dump_json(run/"freeze_manifest.json",frozen)
    base.dump_json(run/"frozen_code_provenance.json",provenance)
    env=base.environment_info();env["method_order_seed"]=protocol["order_seed"]
    base.dump_json(run/"environment.json",env)
    event(run,"code_model_protocol_frozen",**frozen)
    print(f"独立验证运行目录：{run}",flush=True)
    print("代码/模型/协议已封存，现在才开始生成预定六组数据。",flush=True)
    datasets={};data_hashes={}
    for case in protocol["datasets"]:
        gen_case,cfg=generation_spec(case,protocol)
        data,stats=generate_dataset(gen_case,cfg)
        assert len(data.ids)==10032
        for array in (data.ids,data.lon,data.lat):array.flags.writeable=False
        folder=run/"datasets"/case["name"]
        save_dataset(folder,data,gen_case,cfg,stats)
        base.dump_json(folder/"validation_case.json",case)
        datasets[case["name"]]=data
        data_hashes[case["name"]]=sha(folder/"data"/"simulated_points.csv")
        event(run,"dataset_generated",dataset=case["name"],seed=case["seed"],counts=stats,
              input_sha256=data_hashes[case["name"]])
    base.dump_json(run/"data_hashes.json",data_hashes)
    choices=save_prerecommendations(run,datasets,model,protocol)
    print("54条推荐已全部保存并锁定；此后才进行正确性和性能检查。",flush=True)
    checks,accuracy_oracles=preflight(run,datasets,model,protocol,choices)
    records,traces=benchmark(run,datasets,model,protocol,choices,accuracy_oracles)
    base.write_csv(run/"formal_trials.csv",records)
    with (run/"formal_choice_traces.jsonl").open("w",encoding="utf-8") as stream:
        for trace in traces:stream.write(json.dumps(trace,ensure_ascii=False)+"\n")
    summary,paired,comparisons,oracles,regrets=summarize(records,protocol)
    diagnostics,features,pred_errors=posthoc_diagnostics(run,datasets,model,protocol,choices,summary,oracles)
    for filename,rows in (("summary.csv",summary),("paired_differences.csv",paired),("comparisons.csv",comparisons),
                          ("posthoc_oracle.csv",oracles),("strategy_regret.csv",regrets),
                          ("posthoc_diagnostics.csv",diagnostics),("sample_feature_errors.csv",features),
                          ("fixed_prediction_errors.csv",pred_errors)):
        base.write_csv(run/filename,rows)
    write_results(run,summary,comparisons,oracles,regrets,diagnostics,features,protocol)
    triggers=[r for r in diagnostics if r["policy"]=="sample_four" and r["new_resolution_selected"]]
    base.dump_json(run/"results.json",{"independent_datasets":6,"scenarios":18,"repeats":7,
        "formal_records":len(records),"correctness":checks,"summary":summary,"comparisons":comparisons,
        "posthoc_oracle":oracles,"diagnostics":diagnostics,"new_resolution_triggers":triggers,
        "no_trigger_statement":protocol["no_trigger_statement"] if not triggers else None})
    assert PROTOCOL.read_bytes()==protocol_bytes and source.read_bytes()==model_bytes
    assert model==json.loads(model_bytes)
    assert all(sha(ROOT/f)==digest for f,digest in hashes.items())
    assert all(sha(ROOT/p)==digest for p,digest in protected.items())
    assert sha(run/"prerecommendations.jsonl")==json.loads((run/"prerecommendations_lock.json").read_text("utf-8"))["sha256"]
    event(run,"validation_completed_no_refit",formal_records=len(records),new_resolution_triggers=len(triggers))
    base.dump_json(run/"status.json",{"status":"completed_independent_validation","formal_records":len(records),
        "prerecommendations":54,"datasets":6,"scenarios":18,"old_files_unchanged":True,
        "code_model_protocol_unchanged":True,"additional_seeds_or_experiments":False,
        "new_resolution_triggers":len(triggers)})
    base.dump_json(run/"sha256_manifest.json",{p.relative_to(run).as_posix():sha(p) for p in run.rglob("*") if p.is_file()})
    (ROOT/"LATEST_INDEPENDENT_VALIDATION.txt").write_text(run.relative_to(ROOT).as_posix()+"\n",encoding="utf-8")
    for strategy,control in protocol["comparisons"]:
        rows=[r for r in comparisons if r["strategy"]==strategy and r["control"]==control]
        worst=max(rows,key=lambda r:r["worst_paired_extra_ms"])
        print(f"{strategy}相对{control}: 18场景中位数净节省为正{sum(r['median_difference_saving_ms']>0 for r in rows)}项；"
              f"最坏单轮额外{worst['worst_paired_extra_ms']:.4f}ms，{worst['dataset']}/K{worst['K']}。",flush=True)
    print(f"新增档位触发场景数：{len(triggers)}。验证结束，未调参或追加实验。",flush=True)


if __name__=="__main__":main()
