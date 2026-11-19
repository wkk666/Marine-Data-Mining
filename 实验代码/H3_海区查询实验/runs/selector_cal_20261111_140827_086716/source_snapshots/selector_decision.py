"""纯计算的冻结选择规则，不读取数据、不构建索引、不执行查询或计时。

成本单位为毫秒。ABS_EQUALITY_TOL_MS=1e-12 只消除浮点等式附近的
舍入误差；它不是 2% 保守门槛的替代。严格门槛的等式点不允许切换。

模型结构：
    {"f_min": ..., "f_max": ..., "methods": {
        "direct": {"prepare_ms": P, "a_ms": a, "b_ms": b,
                   "mean_query_ms": ...}, ...}}

mean_query_ms 是校准记录，本模块不在 P + K(a + bf) 中使用它。
"""
from __future__ import annotations

import itertools
import math
from collections.abc import Mapping


METHODS = ("direct", "r6", "r7", "r8")
RAW_TIE_ORDER = ("direct", "r7", "r6", "r8")
EXTRA_TIE_ORDER = ("r6", "r8")
ABS_EQUALITY_TOL_MS = 1e-12
ROOT_MERGE_TOL_F = 1e-12
BASE_RATIO = 1.02
EXTRA_RATIO = 0.98


def _finite_float(value, name):
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError(f"{name} must be a finite number") from exc
    if not math.isfinite(result):
        raise ValueError(f"{name} must be a finite number")
    return result


def _model_parts(model, K):
    if not isinstance(model, Mapping):
        raise ValueError("model must be a mapping")
    f_min = _finite_float(model["f_min"], "f_min")
    f_max = _finite_float(model["f_max"], "f_max")
    if f_max < f_min:
        raise ValueError("f_max must be >= f_min")
    k = _finite_float(K, "K")
    if k < 0:
        raise ValueError("K must be nonnegative")
    lines = {}
    for method in METHODS:
        row = model["methods"][method]
        preparation = _finite_float(row["prepare_ms"], f"{method}.prepare_ms")
        a = _finite_float(row["a_ms"], f"{method}.a_ms")
        b = _finite_float(row["b_ms"], f"{method}.b_ms")
        intercept = preparation + k * a
        slope = k * b
        if not math.isfinite(intercept) or not math.isfinite(slope):
            raise ValueError(f"nonfinite total-cost coefficients for {method}")
        lines[method] = (intercept, slope)
    return f_min, f_max, lines


def predict_costs(model, f, K):
    """返回四方法的 P + K(a + bf)，f 先裁剪到校准范围。

    严格遵循冻结表达式，不另对 a+bf 做非负截断，也不混入选择费用。
    """
    f_min, f_max, lines = _model_parts(model, K)
    feature = min(f_max, max(f_min, _finite_float(f, "f")))
    return {method: _finite_float(intercept + slope * feature, f"cost[{method}]")
            for method, (intercept, slope) in lines.items()}


def _argmin_with_ties(costs, priority):
    winner = priority[0]
    for method in priority[1:]:
        if costs[method] - costs[winner] < -ABS_EQUALITY_TOL_MS:
            winner = method
    return winner


def decide(costs, allow_extra: bool):
    """先决定 D/r7 基准，再按严格新增门槛决定是否允许 r6/r8。

    CD <= 1.02*C7 时基准为 direct；否则为 r7。
    新增候选必须严格低于 0.98*CB，等式点不合格。
    两个新增候选同价时选 r6。raw_argmin 的同价优先级为
    direct、r7、r6、r8；raw_argmin 只用于解释记录。

    extra_qualified 始终报告实际合格的方法，即使 allow_extra=False。
    """
    values = {method: _finite_float(costs[method], f"cost[{method}]") for method in METHODS}
    base_delta = values["direct"] - BASE_RATIO * values["r7"]
    base = "direct" if base_delta <= ABS_EQUALITY_TOL_MS else "r7"
    threshold = EXTRA_RATIO * values[base]
    qualified = [method for method in EXTRA_TIE_ORDER
                 if values[method] - threshold < -ABS_EQUALITY_TOL_MS]
    selected = _argmin_with_ties(values, qualified) if allow_extra and qualified else base
    raw = _argmin_with_ties(values, RAW_TIE_ORDER)
    return {"base": base, "selected": selected, "raw_argmin": raw,
            "allow_extra": bool(allow_extra), "extra_qualified": qualified,
            "extra_possible": bool(qualified), "extra_selected": selected in EXTRA_TIE_ORDER,
            "extra_threshold_ms": threshold, "base_cost_ms": values[base],
            "selected_cost_ms": values[selected], "raw_min_cost_ms": values[raw]}


def _state_at(model, feature, K):
    costs = predict_costs(model, feature, K)
    sample2 = decide(costs, allow_extra=False)
    sample4 = decide(costs, allow_extra=True)
    return {"costs_ms": costs, "sample2": sample2, "sample4": sample4,
            "raw_argmin": sample4["raw_argmin"],
            "extra_possible": sample4["extra_possible"]}


def decision_intervals(model, K):
    """解析全部线性交点，返回开区间、边界点和 clip 后的恒定尾部。

    主区间是 [f_min, f_max]，开区间内的 sample2/sample4 给出代表点决策。
    两个 clipped_tails 分别表示 f<f_min 和 f>f_max，决策等于相应端点。
    每个边界点单独执行严格门槛，不能由邻接开区间的标签替代。

    所有交点直接解一元线性方程，不通过网格扫描寻找。恒等的方程另列，
    平行但不相交的方程记 root=None。1e-12 的 root 合并容差只处理数值上
    重合的交点；阈值附近小于浮点等式容差的区域不应作实质性能解释。
    """
    f_min, f_max, lines = _model_parts(model, K)
    equations = [("base:direct=1.02*r7", "direct", "r7", BASE_RATIO)]
    for method in EXTRA_TIE_ORDER:
        for base in ("direct", "r7"):
            equations.append((f"extra:{method}=0.98*{base}", method, base, EXTRA_RATIO))
    equations.append(("extra_tie:r6=r8", "r6", "r8", 1.0))
    for left, right in itertools.combinations(METHODS, 2):
        equations.append((f"raw:{left}={right}", left, right, 1.0))

    equation_roots = []
    raw_roots = [(f_min, "domain_min"), (f_max, "domain_max")]
    coincident = []
    for label, left, right, ratio in equations:
        intercept = lines[left][0] - ratio * lines[right][0]
        slope = lines[left][1] - ratio * lines[right][1]
        if slope == 0.0:
            same_line = abs(intercept) <= ABS_EQUALITY_TOL_MS
            equation_roots.append({"equation": label, "root": None, "coincident": same_line,
                                   "in_domain": same_line})
            if same_line:
                coincident.append(label)
            continue
        root = -intercept / slope
        in_domain = math.isfinite(root) and f_min - ROOT_MERGE_TOL_F <= root <= f_max + ROOT_MERGE_TOL_F
        equation_roots.append({"equation": label, "root": root if math.isfinite(root) else None,
                               "coincident": False, "in_domain": in_domain})
        if in_domain:
            if abs(root - f_min) <= ROOT_MERGE_TOL_F:
                root = f_min
            elif abs(root - f_max) <= ROOT_MERGE_TOL_F:
                root = f_max
            raw_roots.append((root, label))

    raw_roots.sort(key=lambda item: item[0])
    groups = []
    for root, label in raw_roots:
        if groups and abs(root - groups[-1]["f"]) <= ROOT_MERGE_TOL_F:
            groups[-1]["equations"].append(label)
            # 端点必须保持传入的原始值，不被附近交点替代。
            if label == "domain_min":
                groups[-1]["f"] = f_min
            elif label == "domain_max":
                groups[-1]["f"] = f_max
        else:
            groups.append({"f": root, "equations": [label]})

    boundaries = [{**group, **_state_at(model, group["f"], K)} for group in groups]
    intervals = []
    for left, right in zip(groups, groups[1:]):
        midpoint = left["f"] + (right["f"] - left["f"]) / 2.0
        intervals.append({"left": left["f"], "right": right["f"], "left_closed": False,
                          "right_closed": False, "representative_f": midpoint,
                          **_state_at(model, midpoint, K)})
    return {"domain": {"f_min": f_min, "f_max": f_max},
            "cost_lines_ms": {method: {"intercept": pair[0], "slope": pair[1]}
                              for method, pair in lines.items()},
            "equation_roots": equation_roots, "coincident_equations": coincident,
            "boundaries": boundaries, "open_intervals": intervals,
            "clipped_tails": {
                "left": {"condition": "f < f_min", "boundary_f": f_min, **_state_at(model, f_min, K)},
                "right": {"condition": "f > f_max", "boundary_f": f_max, **_state_at(model, f_max, K)}},
            "absolute_equality_tolerance_ms": ABS_EQUALITY_TOL_MS,
            "root_merge_tolerance_f": ROOT_MERGE_TOL_F,
            "raw_tie_order": list(RAW_TIE_ORDER), "extra_tie_order": list(EXTRA_TIE_ORDER)}
