"""统一共同几何、方法专属准备、窄输入选择接口。冻结的查询函数保持原样。"""
from dataclasses import dataclass
import time

import run_experiment as base
from bbox_ablation import query_h3_bbox
from selector_decision import predict_costs, decide

np = base.np
METHODS = ("direct", "r6", "r7", "r8")


@dataclass(frozen=True)
class RawCoordinates:
    longitude: object
    latitude: object

    @property
    def count(self):
        return len(self.longitude)


@dataclass(frozen=True)
class CommonGeometry:
    region: object
    bounds: tuple


def prepare_common():
    region = base.make_region()
    return CommonGeometry(region, region.bounds)


def prepare_specific(method, data, common):
    start = time.perf_counter_ns()
    original_name = "direct" if method == "direct" else "h3_" + method
    state = base.State(original_name, common.region, common.bounds)
    metrics = {"point_index_ns": 0, "point_encoding_ns": 0, "index_grouping_ns": 0,
               "cover_generation_ns": 0, "cell_classification_ns": 0}
    if method != "direct":
        index_start = time.perf_counter_ns()
        state.encoded, state.index, encode, group = base.build_point_index(data, int(method[1:]))
        index_end = time.perf_counter_ns()
        state.enumerated_cells = base.enumerate_cover(int(method[1:]))
        cover_end = time.perf_counter_ns()
        state.cover, state.internal, state.border, state.cell_geometries = base.classify_cells(
            common.region, state.enumerated_cells)
        classify_end = time.perf_counter_ns()
        metrics.update(point_index_ns=index_end - index_start, point_encoding_ns=encode,
                       index_grouping_ns=group, cover_generation_ns=cover_end - index_end,
                       cell_classification_ns=classify_end - cover_end)
    metrics["specific_prepare_ns"] = time.perf_counter_ns() - start
    assert state.region is common.region and state.bounds is common.bounds
    return state, metrics


def query_function(method):
    return base.query if method == "direct" else query_h3_bbox


def r7_feature(raw, common, sample_size=200, sample_seed=73001, diagnostic=False):
    # RawCoordinates没有ID、标签或全量索引；局部类别缓存只在这次特征提取中使用。
    if not isinstance(raw, RawCoordinates):
        raise TypeError("selector accepts RawCoordinates only")
    if diagnostic:
        t0 = time.perf_counter_ns()
    rng = np.random.Generator(np.random.PCG64(sample_seed))
    rows = rng.choice(raw.count, sample_size, replace=False)
    x, y = raw.longitude[rows], raw.latitude[rows]
    if diagnostic:
        t1 = time.perf_counter_ns()
    cells = np.fromiter((base.h3.latlng_to_cell(float(lat), float(lon), 7) for lon, lat in zip(x, y)),
                        dtype=np.uint64, count=sample_size)
    if diagnostic:
        t2 = time.perf_counter_ns()
    unique, inverse = np.unique(cells, return_inverse=True)
    if diagnostic:
        t3 = time.perf_counter_ns()
    _, _, border, _ = base.classify_cells(common.region, tuple(map(int, unique)))
    if diagnostic:
        t4 = time.perf_counter_ns()
    # 分类去重，计数必须通过inverse回到每个样本点。
    unique_is_border = np.isin(unique, np.asarray(border, dtype=np.uint64))
    xmin, ymin, xmax, ymax = common.bounds
    qualifies = unique_is_border[inverse] & (x >= xmin) & (x <= xmax) & (y >= ymin) & (y <= ymax)
    numerator = int(np.count_nonzero(qualifies))
    result = {"f_raw": numerator / sample_size, "sample_numerator": numerator,
              "sample_size": sample_size, "sample_rows": rows.tolist(), "unique_sample_cells": len(unique)}
    if diagnostic:
        t5 = time.perf_counter_ns()
        result["diagnostic_ns"] = {"sampling_and_gather_ns": t1 - t0, "r7_encoding_ns": t2 - t1,
                                   "sample_dedup_ns": t3 - t2, "local_cell_classification_ns": t4 - t3,
                                   "point_weighted_count_and_trace_ns": t5 - t4}
    return result


def choose(raw, common, k, model, strategy, diagnostic=False):
    if strategy not in ("k_only", "sample_two", "sample_four"):
        raise ValueError(strategy)
    if not isinstance(raw, RawCoordinates):
        raise TypeError("selector accepts raw coordinates, never a full prepared State")
    if k == 1:
        return {"selected": "direct", "base": "direct", "raw_argmin": None,
                "reason": "K1_shortcut_no_sampling", "f_raw": None, "f_used": None,
                "sample_rows": [], "sample_numerator": None, "sample_size": 0,
                "costs": {}, "extra_qualified": []}
    if diagnostic:
        start = time.perf_counter_ns()
    if strategy == "k_only":
        feature = {"f_raw": None, "sample_rows": [], "sample_numerator": None, "sample_size": 0}
        if diagnostic:
            before_prediction = time.perf_counter_ns()
        costs = {m: model["methods"][m]["prepare_ms"] + k * model["methods"][m]["mean_query_ms"] for m in METHODS}
        f_used = None
    else:
        feature = r7_feature(raw, common, model["sample_size"], model["sample_seed"], diagnostic)
        if diagnostic:
            before_prediction = time.perf_counter_ns()
        costs = predict_costs(model, feature["f_raw"], k)
        f_used = min(model["f_max"], max(model["f_min"], feature["f_raw"]))
    if diagnostic:
        before_decision = time.perf_counter_ns()
    decision = decide(costs, allow_extra=strategy == "sample_four")
    output = {**feature, **decision, "f_used": f_used, "costs": costs}
    if diagnostic:
        end = time.perf_counter_ns()
        diag = dict(feature.get("diagnostic_ns", {}))
        diag.update(feature_total_ns=before_prediction - start, prediction_ns=before_decision - before_prediction,
                    final_decision_and_trace_ns=end - before_decision)
        output["diagnostic_ns"] = diag
    return output


def execute_task(data, k, model=None, strategy=None, fixed_method=None):
    """可供后续端到端计时：每个任务只创建一次共同几何，查询不缓存答案。

    本轮不以分项中位数冒充此流程的总体性能。
    """
    if (strategy is None) == (fixed_method is None):
        raise ValueError("specify exactly one strategy or fixed method")
    common = prepare_common()
    if strategy is None:
        choice = {"selected": fixed_method, "reason": "fixed_method"}
    else:
        choice = choose(RawCoordinates(data.lon, data.lat), common, k, model, strategy)
    state, _ = prepare_specific(choice["selected"], data, common)
    query = query_function(choice["selected"])
    for _ in range(k):
        ids, _ = query(state, data)
    return ids, choice
