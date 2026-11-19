"""从封存文件绘图；不导入查询/生成/校准模块，不执行实验。"""
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE / ".plot_packages"))
sys.path.insert(0, str(ROOT / ".runtime" / "packages"))
os.environ["MPLCONFIGDIR"] = str(HERE / ".mplconfig")
import csv
import hashlib
import json
import warnings
from importlib import metadata
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager, ticker
from matplotlib.lines import Line2D
from matplotlib.patches import Polygon, Rectangle

V = ROOT / "runs" / "validation_20261111_151150_257869"
C = ROOT / "runs" / "selector_cal_20261111_140827_086716"
FIG = HERE / "图稿"
DATA = HERE / "图表数据"
for folder in (FIG, DATA):
    folder.mkdir(parents=True, exist_ok=True)
FONT = Path("C:/Windows/Fonts/msyh.ttc")
font_manager.fontManager.addfont(str(FONT))
FONT_NAME = font_manager.FontProperties(fname=str(FONT)).get_name()
plt.rcParams.update({
    "font.family": FONT_NAME, "font.size": 13, "axes.titlesize": 14,
    "axes.labelsize": 13, "xtick.labelsize": 11.5, "ytick.labelsize": 12,
    "axes.unicode_minus": True, "svg.fonttype": "path", "savefig.dpi": 300,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.edgecolor": "#657078", "axes.linewidth": .7,
    "grid.color": "#DCE1E4", "grid.linewidth": .6,
    "figure.facecolor": "white", "axes.facecolor": "white",
})
COLORS = {"direct": "#555B61", "r6": "#D48A22", "r7": "#2376A8", "r8": "#8F62A7"}
NAMES = {"direct": "直接查询", "r6": "新版 r6", "r7": "新版 r7", "r8": "新版 r8"}
CASES = ["uniform_3101", "uniform_3102", "moderate_3201", "moderate_3202", "strong_3301", "strong_3302"]
LABELS = ["均匀 3101", "均匀 3102", "中等 3201", "中等 3202", "较强 3301", "较强 3302"]
METHODS = ["direct", "r6", "r7", "r8"]
INPUTS = {}


def read_bytes(path):
    content = path.read_bytes()
    INPUTS[str(path)] = hashlib.sha256(content).hexdigest()
    return content


def load_csv(path):
    return list(csv.DictReader(read_bytes(path).decode("utf-8-sig").splitlines()))


def load_json(path):
    return json.loads(read_bytes(path).decode("utf-8"))


def write_csv(path, rows):
    with path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def save(fig, name):
    # 不用tight裁剪改变图稿尺寸；版式由显式边距控制。
    fig.savefig(FIG / (name + ".png"), dpi=300)
    fig.savefig(FIG / (name + ".svg"))
    plt.close(fig)


def figure1():
    chosen = [CASES[i] for i in (0, 2, 4)]
    titles = ["(a) 均匀：10000 背景点", "(b) 中等：40% 边界带成分", "(c) 较强：80% 边界带成分"]
    fig = plt.figure(figsize=(12.6, 6.8))
    gs = fig.add_gridspec(2, 3, height_ratios=[3.4, 1.15], left=.06, right=.985,
                          top=.86, bottom=.15, wspace=.12, hspace=.62)
    all_data = []
    for j, (case, title) in enumerate(zip(chosen, titles)):
        folder = V / "datasets" / case / "data"
        rows = load_csv(folder / "simulated_points.csv")
        sea = load_json(folder / "simulated_sea.geojson")
        if sea["type"] == "Feature":
            sea = sea["geometry"]
        verts = np.asarray(sea["coordinates"][0], float)
        x = np.array([float(r["longitude"]) for r in rows])
        y = np.array([float(r["latitude"]) for r in rows])
        probes = np.array([bool(r["probe_expected_inside"]) for r in rows])
        assert len(rows) == 10032 and probes.sum() == 32
        all_data.append((rows, x, y, probes))
        ax = fig.add_subplot(gs[0, j])
        ax.add_patch(Polygon(verts, closed=True, facecolor="#E6EEE9", edgecolor="none", zorder=0))
        ax.scatter(x[~probes], y[~probes], s=2.7, alpha=.24, c="#397DA2", linewidths=0, zorder=1)
        ax.plot(verts[:, 0], verts[:, 1], c="#263E39", lw=1.6, zorder=2)
        ax.scatter(x[probes], y[probes], s=25, facecolors="none", edgecolors="#BE3F43", lw=1.2, zorder=3)
        ax.set(xlim=(121.875, 122.375), ylim=(29.875, 30.375), aspect="equal", xlabel="经度 / °E")
        if j == 0:
            ax.set_ylabel("纬度 / °N")
        else:
            ax.tick_params(labelleft=False)
        ax.xaxis.set_major_locator(ticker.MultipleLocator(.2))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(.2))
        ax.xaxis.set_major_formatter(ticker.FormatStrFormatter("%.1f"))
        ax.yaxis.set_major_formatter(ticker.FormatStrFormatter("%.1f"))
        ax.set_title(title, fontsize=13, pad=10)
        ax.text(.03, .03, case.split("_")[-1], transform=ax.transAxes, fontsize=11,
                bbox=dict(facecolor="white", alpha=.9, edgecolor="none"))
    left = fig.add_subplot(gs[1, :2]); left.axis("off")
    handles = [Line2D([], [], marker=".", color="#397DA2", ls="", markersize=9, label="模拟船位记录（全部绘制）"),
               Line2D([], [], color="#263E39", lw=1.6, label="模拟海区边界"),
               Line2D([], [], marker="o", markerfacecolor="none", markeredgecolor="#BE3F43", ls="", label="32 条边界测试记录")]
    left.legend(handles=handles, loc="upper left", frameon=False, fontsize=12, ncol=1)
    left.text(.02, -.04, "主图尺度下测试点局部重叠；右侧放大同一条边上的三点。\n点坐标未抽样、未移动；放大图仅改变观察范围。", fontsize=11, va="top")
    inset = fig.add_subplot(gs[1, 2])
    rows, x, y, probes = all_data[0]
    cats = ["near_outside_0", "on_edge_0", "near_inside_0"]
    marks = ["x", "o", "^"]
    for category, mark in zip(cats, marks):
        i = next(i for i, r in enumerate(rows) if r["category"] == category)
        inset.scatter([x[i]], [y[i]], s=40, marker=mark, c="#BE3F43", zorder=3)
    inset.axhline(30, c="#263E39", lw=.8)
    inset.set(xlim=(122.12499985, 122.12500015), ylim=(29.99999985, 30.00000015))
    inset.set_xticks([122.125], ["经度 122.125°E"])
    inset.set_yticks([29.9999999, 30, 30.0000001], ["边外", "边界上", "边内"])
    inset.tick_params(labelsize=10)
    inset.set_title("边 0 局部放大（原坐标）", fontsize=11)
    fig.suptitle("图 1  固定模拟海区与三种模拟船位分布", y=.965, fontsize=17)
    fig.text(.5, .025, "每组均为 10000 条随机记录 + 原 32 条测试记录；生成成分比例不等于实测近边界比例。", ha="center", fontsize=11)
    save(fig, "图1_海区与点分布")


def figure2(summary):
    lookup = {(r["dataset"], int(r["K"]), r["scheme"]): r for r in summary}
    fig, axes = plt.subplots(3, 2, figsize=(12.0, 10.1), sharey=True)
    fig.subplots_adjust(left=.09, right=.985, top=.88, bottom=.12, hspace=.36, wspace=.15)
    markers = ["o", "s", "D", "^"]
    offsets = [-.24, -.08, .08, .24]
    plot_rows = []
    for i, (case, label) in enumerate(zip(CASES, LABELS)):
        ax = axes.flat[i]
        for m, offset, marker in zip(METHODS, offsets, markers):
            for pos, k in enumerate([1, 100, 1000]):
                r = lookup[case, k, "fixed_" + m]
                med, q25, q75 = (float(r[t]) for t in ["median_ms", "q25_ms", "q75_ms"])
                ax.errorbar(pos + offset, med, yerr=[[med - q25], [q75 - med]],
                            fmt=marker, ms=6, color=COLORS[m], elinewidth=1.5, capsize=3, zorder=3)
                plot_rows.append({"dataset":case, "K":k, "method":m, "median_ms":med, "q25_ms":q25, "q75_ms":q75})
        ax.set_yscale("log")
        ax.set_ylim(.14, 700)
        ax.set_xlim(-.5, 2.5)
        ax.set_xticks([0, 1, 2], ["1", "100", "1000"])
        ax.set_xlabel("重新执行次数 K（离散实测条件）", fontsize=12)
        ax.grid(axis="y", which="major", alpha=.8)
        ax.set_title(f"({chr(97+i)}) {label}", loc="left")
        ax.set_yticks([.2, 1, 10, 100, 500], ["0.2", "1", "10", "100", "500"])
        if i % 2 == 0:
            ax.set_ylabel("完整总时间 / ms（对数坐标）")
        if i >= 4:
            d = float(lookup[case, 1000, "fixed_r7"]["median_ms"]) - float(lookup[case, 1000, "fixed_r6"]["median_ms"])
            ax.text(.035, .78, f"K=1000：r6 中位数仅比 r7 少 {d:.4f} ms", transform=ax.transAxes, fontsize=11,
                    bbox=dict(facecolor="white", edgecolor="none", alpha=.9))
    handles = [Line2D([], [], color=COLORS[m], marker=mark, ls="", label=NAMES[m], markersize=7)
               for m, mark in zip(METHODS, markers)]
    fig.legend(handles=handles, loc="upper center", bbox_to_anchor=(.5, .95), ncol=4, frameon=False)
    fig.suptitle("图 2  四种固定方法的完整总成本", y=.985, fontsize=17)
    fig.text(.5, .035, "一次共同准备 + 方法专属准备 + K 次正常查询；点为中位数，误差棒为第25—75百分位（非置信区间）。\n仅固定方法；未绘制样本策略，无连续性能曲线。", ha="center", fontsize=11.5)
    write_csv(DATA / "图2_固定方法绘图数据.csv", plot_rows)
    save(fig, "图2_固定方法完整成本")


PAIRS = [("sample_two", "k_only", "样本二选一相对 K-only"),
         ("sample_four", "sample_two", "样本四选一相对样本二选一"),
         ("sample_four", "fixed_direct", "样本四选一相对固定直接查询"),
         ("sample_four", "fixed_r7", "样本四选一相对固定 r7")]


def figure3(paired):
    fig, axes = plt.subplots(4, 3, figsize=(12.8, 13.5))
    fig.subplots_adjust(left=.11, right=.987, bottom=.10, top=.91, wspace=.15, hspace=.46)
    stats = []
    for row, (strategy, control, title) in enumerate(PAIRS):
        for col, k in enumerate([1, 100, 1000]):
            ax = axes[row, col]
            largest = .06
            for i, case in enumerate(CASES):
                records = sorted([r for r in paired if r["dataset"]==case and int(r["K"])==k
                                  and r["strategy"]==strategy and r["control"]==control], key=lambda r:int(r["trial"]))
                assert len(records)==7
                values = np.array([float(r["net_saving_ms"]) for r in records])
                q25, med, q75 = np.quantile(values, [.25, .5, .75], method="linear")
                largest = max(largest, max(abs(values)))
                # 纵向错位仅区分轮次，横坐标保留原始数值；没有随机抖动。
                ax.scatter(values, i + np.linspace(-.14, .14, 7), s=14, facecolors="none", edgecolors="#858B90", lw=.65, zorder=2)
                ax.errorbar(med, i, xerr=[[med-q25], [q75-med]], fmt="D", ms=5,
                            color="#2376A8", elinewidth=2.2, capsize=4, zorder=4)
                stats.append({"dataset":case,"K":k,"strategy":strategy,"control":control,
                              "paired_median_ms":med,"q25_ms":q25,"q75_ms":q75,
                              "min_ms":values.min(),"max_ms":values.max(),
                              **{f"trial_{t}_ms":v for t,v in enumerate(values,1)}})
            bound = largest * 1.25
            ax.set_xscale("symlog", linthresh=.1, linscale=.7, base=10)
            ax.set_xlim(-bound, bound)
            ticks = [-.1,-.05,0,.05,.1] if bound<=.3 else [-100,-10,-1,-.1,0,.1,1,10,100]
            ticks = [t for t in ticks if -bound <= t <= bound]
            ax.set_xticks(ticks)
            ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:g}".replace("-", "−")))
            ax.axvline(0, color="#28373F", lw=1, zorder=1)
            ax.grid(axis="x", alpha=.8)
            ax.set_yticks(range(6), LABELS if col==0 else [""]*6)
            ax.set_ylim(5.55,-.6)
            ax.tick_params(axis="x", labelsize=10.5)
            ax.set_title(f"K = {k}", fontsize=13, pad=7)
            if col==0:
                bbox=ax.get_position()
                fig.text(bbox.x0, bbox.y1+.024, f"({chr(97+row)}) {title}", fontsize=14, weight="bold")
    fig.suptitle("图 3  选择策略的同轮净节省与原始波动", y=.985, fontsize=17)
    handles=[Line2D([],[],marker="o",mfc="none",mec="#858B90",ls="",label="七轮原始差值",ms=5),
             Line2D([],[],color="#2376A8",marker="D",lw=2,label="差值中位数及 IQR",ms=5)]
    fig.legend(handles=handles,loc="upper center",bbox_to_anchor=(.5,.96),ncol=2,frameon=False,fontsize=12)
    fig.text(.5,.025,"净节省 / ms = 同轮对照 − 策略；正值为策略更快。各面板横轴为对称对数，|净节省|≤0.1 ms 内线性。\n每场景七个原始差值均保留；IQR 为第25—75百分位，不是最坏损失或置信区间。",ha="center",fontsize=11.5)
    write_csv(DATA / "图3_同轮净节省绘图数据.csv", stats)
    write_csv(DATA / "图3_原始配对差.csv", paired)
    save(fig,"图3_策略净节省与波动")


def merged_bars(intervals, k, policy, lo, hi):
    pieces = sorted([(max(lo,float(r["left"])*100),min(hi,float(r["right"])*100),r[policy])
                     for r in intervals if int(r["K"])==k and r["kind"] in ["open","clip_left","clip_right"]
                     and float(r["right"])*100>lo and float(r["left"])*100<hi])
    merged=[]
    for left,right,method in pieces:
        if merged and method==merged[-1][2] and abs(left-merged[-1][1])<1e-8:
            merged[-1]=(merged[-1][0],right,method)
        else:
            merged.append((left,right,method))
    return merged


def plot_feature_panel(ax, features, intervals, model, k, zoom=False):
    lo,hi=(49.5,54.7) if zoom else (4.0,55.5)
    fmin,fmax=model["f_min"]*100,model["f_max"]*100
    if lo<fmin: ax.axvspan(lo,fmin,color="#EFF0F1",zorder=0)
    if hi>fmax: ax.axvspan(fmax,hi,color="#EFF0F1",zorder=0)
    for bound in [fmin,fmax]:
        if lo<bound<hi: ax.axvline(bound,color="#939BA0",ls="--",lw=.8,zorder=1)
    for y, policy in [(-1.45,"sample_two"),(-.85,"sample_four")]:
        for left,right,method in merged_bars(intervals,k,policy,lo,hi):
            ax.add_patch(Rectangle((left,y-.18),right-left,.36,facecolor=COLORS[method],alpha=.88,zorder=2))
            if right-left > (hi-lo)*.12:
                ax.text((left+right)/2,y,"直接" if method=="direct" else method,ha="center",va="center",color="white",fontsize=10)
    selected=features[-2:] if zoom else features
    for y,r in enumerate(selected):
        actual, raw, used = [100*float(r[key]) for key in ["exact_f","sample_f_raw","sample_f_used"]]
        ax.plot([actual,raw,used],[y-.15,y,y+.15],c="#B2B6B8",lw=.8,zorder=2)
        ax.scatter(actual,y-.15,marker="D",c="#252A2F",s=35,zorder=4)
        ax.scatter(raw,y,marker="o",facecolors="white",edgecolors="#168475",s=43,lw=1.4,zorder=5)
        ax.scatter(used,y+.15,marker="s",facecolors="none",edgecolors="#BA4465",s=38,lw=1.4,zorder=5)
    ax.set(xlim=(lo,hi),ylim=(len(selected)-.45,-1.9),xlabel="r7 筛后精判比例 / %")
    labs=LABELS[-2:] if zoom else LABELS
    ax.set_yticks([-1.45,-.85]+list(range(len(selected))),["二选一","四选一"]+labs)
    ax.grid(axis="x",alpha=.6)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=11.5)
    if zoom:
        ax.set_xticks([50,51,52,53,54])
        ax.set_title("(c) K=1000 局部放大：保留原比例位置",fontsize=13,loc="left")
    else:
        ax.set_xticks([5,10,20,30,40,50,55])
        ax.set_title(f"({'a' if k==100 else 'b'}) K = {k}：整体范围",fontsize=14,loc="left")


def figure4(features,intervals,model):
    byname={r["dataset"]:r for r in features};features=[byname[n] for n in CASES]
    fig=plt.figure(figsize=(12.6,7.5))
    gs=fig.add_gridspec(2,2,height_ratios=[1.75,1],left=.11,right=.985,top=.88,bottom=.14,wspace=.25,hspace=.42)
    plot_feature_panel(fig.add_subplot(gs[0,0]),features,intervals,model,100)
    plot_feature_panel(fig.add_subplot(gs[0,1]),features,intervals,model,1000)
    leg=fig.add_subplot(gs[1,0]);leg.axis("off")
    handles=[Line2D([],[],marker="D",c="#252A2F",ls="",label="准确比例（事后诊断）",ms=6),
             Line2D([],[],marker="o",mfc="white",mec="#168475",ls="",label="原始 200 点样本比例",ms=7),
             Line2D([],[],marker="s",mfc="none",mec="#BA4465",ls="",label="截断后用于预测的比例",ms=7)]
    leg.legend(handles=handles,loc="upper left",frameon=False,fontsize=12)
    leg.text(.02,.22,"灰底：校准范围外，预测时仍按冻结规则截断。\nK=1000 的 r6 窗口：[50.6416%, 51.6812%)。\n强聚集准确比例已越过窗口上界；K=1 不抽样。",fontsize=11.5,va="top")
    plot_feature_panel(fig.add_subplot(gs[1,1]),features,intervals,model,1000,zoom=True)
    fig.suptitle("图 4  比例估计与冻结决策区间",y=.98,fontsize=17)
    fig.text(.5,.025,"色条为冻结规则的最终选择；准确比例不曾提供给在线选择器。二选一与四选一在全部18场景选择相同。",ha="center",fontsize=11)
    save(fig,"图4_比例与冻结区间")


def main():
    summary=load_csv(V/"summary.csv")
    paired=load_csv(V/"paired_differences.csv")
    features=load_csv(V/"sample_feature_errors.csv")
    intervals=load_csv(C/"decision_intervals.csv")
    model=load_json(V/"online_model.json")
    assert len(summary)==126 and len(paired)==504 and len(features)==6
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        figure1();figure2(summary);figure3(paired);figure4(features,intervals,model)
    warning_text=[str(w.message) for w in caught]
    assert not any("Glyph" in s and "missing" in s for s in warning_text),warning_text
    assert all(hashlib.sha256(Path(path).read_bytes()).hexdigest()==digest for path,digest in INPUTS.items())
    write_csv(DATA/"图4_比例绘图数据.csv",features)
    write_csv(DATA/"图4_冻结区间原记录.csv",intervals)
    provenance={"plot_script_version":"report-figures-v1","python":sys.version,"numpy":np.__version__,
                "matplotlib":matplotlib.__version__,"font":str(FONT),"font_name":FONT_NAME,
                "dependencies":{name:metadata.version(name) for name in
                                ["matplotlib","numpy","Pillow","packaging","python-dateutil",
                                 "contourpy","cycler","fonttools","kiwisolver","pyparsing","six"]},
                "font_sha256":hashlib.sha256(FONT.read_bytes()).hexdigest(),
                "input_sha256":INPUTS,"all_inputs_unchanged":True,
                "query_or_generation_modules_imported":False,"new_random_sampling":False,
                "query_timing_or_fitting_executed":False,"warnings":warning_text,
                "quantile":"numpy.quantile(method='linear')","png_dpi":300,"vector":"SVG; font outlines"}
    (HERE/"绘图来源与环境.json").write_text(json.dumps(provenance,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps({"figures":8,"source_files":len(INPUTS),"warnings":warning_text,"status":"plotted_from_saved_data_only"},ensure_ascii=False))


if __name__=="__main__":
    main()
