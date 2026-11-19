"""只读取封存数据，另导出160 mm正文图。保留原总览图与实验代码。"""
from pathlib import Path
import os
import sys

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
OUT = HERE / "图稿" / "正文版"
OUT.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, str(HERE / ".plot_packages"))
sys.path.insert(0, str(ROOT / ".runtime" / "packages"))
os.environ["MPLCONFIGDIR"] = str(OUT / ".mplconfig")
import csv
import hashlib
import json
import warnings
import xml.etree.ElementTree as ET
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager, ticker
from matplotlib.lines import Line2D
from matplotlib.patches import Polygon, Rectangle

V = ROOT / "runs" / "validation_20261111_151150_257869"
C = ROOT / "runs" / "selector_cal_20261111_140827_086716"
WIDTH_MM = 160
FONT = Path("C:/Windows/Fonts/msyh.ttc")
font_manager.fontManager.addfont(str(FONT))
FONT_NAME = font_manager.FontProperties(fname=str(FONT)).get_name()
plt.rcParams.update({
    "font.family": FONT_NAME, "font.size": 8.5, "axes.titlesize": 9,
    "axes.labelsize": 8.5, "xtick.labelsize": 8, "ytick.labelsize": 8,
    "axes.unicode_minus": True, "svg.fonttype": "path", "savefig.dpi": 300,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.edgecolor": "#657078", "axes.linewidth": .55,
    "grid.color": "#DCE1E4", "grid.linewidth": .5,
    "figure.facecolor": "white", "axes.facecolor": "white",
})
COLORS = {"direct":"#555B61", "r6":"#D48A22", "r7":"#2376A8", "r8":"#8F62A7"}
NAMES = {"direct":"直接查询", "r6":"新版 r6", "r7":"新版 r7", "r8":"新版 r8"}
CASES = ["uniform_3101","uniform_3102","moderate_3201","moderate_3202","strong_3301","strong_3302"]
LABELS = ["均匀 3101","均匀 3102","中等 3201","中等 3202","较强 3301","较强 3302"]
METHODS = ["direct","r6","r7","r8"]
INPUTS = {}
OUTPUTS = []


def read(path):
    b = path.read_bytes()
    INPUTS[path.relative_to(ROOT).as_posix()] = hashlib.sha256(b).hexdigest()
    return b.decode("utf-8-sig")


def csv_rows(path):
    return list(csv.DictReader(read(path).splitlines()))


def json_obj(path):
    return json.loads(read(path))


def fig_mm(height):
    return plt.figure(figsize=(WIDTH_MM/25.4, height/25.4))


def save(fig, name, height, **counts):
    for suffix in ("png","svg"):
        fig.savefig(OUT / f"{name}.{suffix}", dpi=300)
    OUTPUTS.append({"name":name,"width_mm":WIDTH_MM,"height_mm":height,
                    "font_pt":{"body":8.5,"tick":8,"title":11},**counts})
    plt.close(fig)


def figure1():
    height = 105
    fig = fig_mm(height)
    gs = fig.add_gridspec(2,3,height_ratios=[3.2,1],left=.085,right=.99,top=.84,bottom=.15,
                          wspace=.10,hspace=.60)
    chosen = [CASES[i] for i in (0,2,4)]
    titles = ["(a) 均匀 · 3101","(b) 中等 · 3201","(c) 较强 · 3301"]
    arrays=[]
    for j,(case,title) in enumerate(zip(chosen,titles)):
        folder=V/"datasets"/case/"data"
        rows=csv_rows(folder/"simulated_points.csv")
        sea=json_obj(folder/"simulated_sea.geojson")
        if sea["type"]=="Feature":sea=sea["geometry"]
        verts=np.asarray(sea["coordinates"][0],float)
        x=np.array([float(r["longitude"]) for r in rows]);y=np.array([float(r["latitude"]) for r in rows])
        probe=np.array([bool(r["probe_expected_inside"]) for r in rows])
        assert len(rows)==10032 and probe.sum()==32
        arrays.append((rows,x,y,probe))
        ax=fig.add_subplot(gs[0,j])
        ax.add_patch(Polygon(verts,closed=True,facecolor="#E6EEE9",edgecolor="none",zorder=0))
        ax.scatter(x[~probe],y[~probe],s=.45,alpha=.25,c="#397DA2",linewidths=0,zorder=1)
        ax.plot(verts[:,0],verts[:,1],c="#263E39",lw=.8,zorder=2)
        ax.scatter(x[probe],y[probe],s=7,facecolors="none",edgecolors="#BE3F43",lw=.65,zorder=3)
        ax.set(xlim=(121.875,122.375),ylim=(29.875,30.375),aspect="equal",xlabel="经度 / °E")
        ax.set_xticks([122.0,122.2]);ax.set_yticks([30.0,30.2])
        ax.xaxis.set_major_formatter(ticker.FormatStrFormatter("%.1f"));ax.yaxis.set_major_formatter(ticker.FormatStrFormatter("%.1f"))
        if j==0:ax.set_ylabel("纬度 / °N")
        else:ax.tick_params(labelleft=False)
        ax.set_title(title,fontsize=8.5,pad=5)
        ax.text(.04,.04,["0%","40%","80%"][j]+" 边界带成分",transform=ax.transAxes,fontsize=8,
                bbox=dict(facecolor="white",alpha=.92,edgecolor="none"))
    leg=fig.add_subplot(gs[1,:2]);leg.axis("off")
    handles=[Line2D([],[],marker=".",color="#397DA2",ls="",ms=5,label="模拟记录（全部绘制）"),
             Line2D([],[],color="#263E39",lw=.8,label="模拟海区边界"),
             Line2D([],[],marker="o",mfc="none",mec="#BE3F43",ls="",ms=3,label="32条边界测试记录")]
    leg.legend(handles=handles,loc="upper left",frameon=False,fontsize=8,handlelength=1.2,borderpad=0,labelspacing=.35)
    leg.text(0,-.12,"主图测试点局部重叠；放大图不移动坐标。",fontsize=8,va="top")
    ax=fig.add_subplot(gs[1,2]);rows,x,y,_=arrays[0]
    for cat,marker in zip(["near_outside_0","on_edge_0","near_inside_0"],["x","o","^"]):
        i=next(i for i,r in enumerate(rows) if r["category"]==cat)
        ax.scatter(x[i],y[i],s=14,marker=marker,c="#BE3F43",zorder=3)
    ax.axhline(30,c="#263E39",lw=.6)
    ax.set(xlim=(122.12499985,122.12500015),ylim=(29.99999985,30.00000015))
    ax.set_xticks([122.125],["122.125°E"])
    ax.set_yticks([29.9999999,30,30.0000001],["边外","边界","边内"])
    ax.tick_params(labelsize=8)
    ax.set_title("边0原坐标放大",fontsize=8,pad=4)
    fig.suptitle("图1  固定模拟海区与三种模拟船位分布",y=.97,fontsize=11)
    fig.text(.5,.045,"每组10000条随机记录＋32条测试记录；40%/80%是生成成分比例。",ha="center",fontsize=8)
    save(fig,"图1_海区与点分布",height,records_in_main=30096,probe_records_in_main=96)


def figure2(summary):
    height=175
    fig=fig_mm(height)
    axes=fig.subplots(3,2,sharey=True)
    fig.subplots_adjust(left=.115,right=.985,top=.87,bottom=.15,hspace=.45,wspace=.17)
    lookup={(r["dataset"],int(r["K"]),r["scheme"]):r for r in summary}
    markers=["o","s","D","^"]
    for i,(case,label) in enumerate(zip(CASES,LABELS)):
        ax=axes.flat[i]
        for m,off,marker in zip(METHODS,[-.24,-.08,.08,.24],markers):
            for pos,k in enumerate([1,100,1000]):
                r=lookup[case,k,"fixed_"+m]
                med,q25,q75=[float(r[t]) for t in ["median_ms","q25_ms","q75_ms"]]
                ax.errorbar(pos+off,med,yerr=[[med-q25],[q75-med]],fmt=marker,ms=3.4,
                            color=COLORS[m],elinewidth=.9,capsize=1.8,zorder=3)
        ax.set_yscale("log");ax.set_ylim(.14,700);ax.set_xlim(-.5,2.5)
        ax.set_xticks([0,1,2],["1","100","1000"])
        ax.set_yticks([.2,1,10,100,500],["0.2","1","10","100","500"])
        ax.set_xlabel("重新执行次数 K",labelpad=2)
        ax.set_title(f"({chr(97+i)}) {label}",loc="left",pad=5)
        ax.grid(axis="y",alpha=.8)
        if i%2==0:ax.set_ylabel("完整总时间 / ms（对数）",labelpad=3)
        if i>=4:
            diff=float(lookup[case,1000,"fixed_r7"]["median_ms"])-float(lookup[case,1000,"fixed_r6"]["median_ms"])
            ax.text(.02,.77,f"K=1000：r6仅少{diff:.4f} ms",transform=ax.transAxes,fontsize=8,
                    bbox=dict(facecolor="white",alpha=.95,edgecolor="none",pad=1))
    handles=[Line2D([],[],color=COLORS[m],marker=mark,ls="",ms=4,label=NAMES[m]) for m,mark in zip(METHODS,markers)]
    fig.suptitle("图2  四种固定方法的完整总成本",y=.98,fontsize=11)
    fig.legend(handles=handles,loc="upper center",bbox_to_anchor=(.5,.945),ncol=4,frameon=False,
               fontsize=8,handletextpad=.4,columnspacing=1.3)
    fig.text(.5,.04,"一次准备＋K次正常查询；点为中位数，误差棒为IQR（非置信区间）。\n三种K为离散实测条件；只比较固定方法，不含样本选择。",ha="center",fontsize=8)
    save(fig,"图2_固定方法完整成本",height,fixed_statistics=72)


PAIRS=[("sample_two","k_only","图3a_二选一相对Konly","图3a  样本二选一相对 K-only"),
       ("sample_four","sample_two","图3b_四选一相对二选一","图3b  样本四选一相对样本二选一"),
       ("sample_four","fixed_direct","图3c_四选一相对直接","图3c  样本四选一相对固定直接查询"),
       ("sample_four","fixed_r7","图3d_四选一相对r7","图3d  样本四选一相对固定 r7")]


def figure3(paired):
    height=158
    for strategy,control,name,title in PAIRS:
        fig=fig_mm(height);axes=fig.subplots(3,1)
        fig.subplots_adjust(left=.16,right=.98,top=.86,bottom=.13,hspace=.45)
        for ax,k in zip(axes,[1,100,1000]):
            largest=.06
            for i,case in enumerate(CASES):
                records=sorted([r for r in paired if r["dataset"]==case and int(r["K"])==k and
                                r["strategy"]==strategy and r["control"]==control],key=lambda r:int(r["trial"]))
                assert len(records)==7
                values=np.array([float(r["net_saving_ms"]) for r in records])
                q25,med,q75=np.quantile(values,[.25,.5,.75],method="linear")
                largest=max(largest,float(max(abs(values))))
                ax.scatter(values,i+np.linspace(-.15,.15,7),s=8,facecolors="none",edgecolors="#858B90",lw=.5,zorder=2)
                ax.errorbar(med,i,xerr=[[med-q25],[q75-med]],fmt="D",ms=3.5,color="#2376A8",
                            elinewidth=1.3,capsize=2.4,zorder=4)
            bound=largest*1.25
            ax.set_xscale("symlog",linthresh=.1,linscale=.7,base=10)
            ax.set_xlim(-bound,bound)
            ticks=[-.1,-.05,0,.05,.1] if bound<=.3 else [-100,-10,-1,-.1,0,.1,1,10,100]
            ticks=[t for t in ticks if -bound<=t<=bound]
            ax.set_xticks(ticks)
            ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda v,_:f"{v:g}".replace("-","−")))
            ax.axvline(0,c="#28373F",lw=.8,zorder=1);ax.grid(axis="x",alpha=.8)
            ax.set_yticks(range(6),LABELS);ax.set_ylim(5.5,-.6)
            ax.tick_params(axis="y",length=2)
            ax.set_title(f"K = {k}",loc="left",fontsize=9,pad=5)
        handles=[Line2D([],[],marker="o",mfc="none",mec="#858B90",ls="",ms=3,label="七轮原始差值"),
                 Line2D([],[],color="#2376A8",marker="D",lw=1.2,ms=3,label="同轮差值中位数及 IQR")]
        fig.suptitle(title,y=.978,fontsize=11)
        fig.legend(handles=handles,loc="upper center",bbox_to_anchor=(.5,.943),ncol=2,frameon=False,
                   fontsize=8.5,handletextpad=.5,columnspacing=2)
        fig.text(.5,.028,"净节省 / ms = 同轮对照 − 策略；正数为策略更快。\n横轴对称对数，|净节省|≤0.1 ms内线性；七个原始值均保留。\nIQR为第25—75百分位（非置信区间）；各K横轴范围不同。",ha="center",fontsize=8)
        save(fig,name,height,scenarios=18,raw_values=126,medians=18,iqr_bars=18)


def merged_bars(intervals,k,policy,lo,hi):
    pieces=sorted([(max(lo,float(r["left"])*100),min(hi,float(r["right"])*100),r[policy])
                   for r in intervals if int(r["K"])==k and r["kind"] in ["open","clip_left","clip_right"]
                   and float(r["right"])*100>lo and float(r["left"])*100<hi])
    result=[]
    for left,right,method in pieces:
        if result and method==result[-1][2] and abs(left-result[-1][1])<1e-8:
            result[-1]=(result[-1][0],right,method)
        else:result.append((left,right,method))
    return result


def feature_panel(ax,features,intervals,model,k,zoom=False):
    lo,hi=(49.5,54.7) if zoom else (4.0,55.5)
    fmin,fmax=model["f_min"]*100,model["f_max"]*100
    if lo<fmin:ax.axvspan(lo,fmin,color="#EFF0F1",zorder=0)
    if hi>fmax:ax.axvspan(fmax,hi,color="#EFF0F1",zorder=0)
    for bound in [fmin,fmax]:
        if lo<bound<hi:ax.axvline(bound,c="#939BA0",ls="--",lw=.65,zorder=1)
    for y,policy in [(-1.45,"sample_two"),(-.85,"sample_four")]:
        for left,right,method in merged_bars(intervals,k,policy,lo,hi):
            ax.add_patch(Rectangle((left,y-.18),right-left,.36,facecolor=COLORS[method],alpha=.88,zorder=2))
            if right-left>(hi-lo)*.12:
                ax.text((left+right)/2,y,"直接" if method=="direct" else method,ha="center",va="center",c="white",fontsize=8)
    chosen=features[-2:] if zoom else features
    for y,r in enumerate(chosen):
        actual,raw,used=[100*float(r[key]) for key in ["exact_f","sample_f_raw","sample_f_used"]]
        ax.plot([actual,raw,used],[y-.15,y,y+.15],c="#B2B6B8",lw=.6,zorder=2)
        ax.scatter(actual,y-.15,marker="D",c="#252A2F",s=13,zorder=4)
        ax.scatter(raw,y,marker="o",facecolors="white",edgecolors="#168475",s=17,lw=.85,zorder=5)
        ax.scatter(used,y+.15,marker="s",facecolors="none",edgecolors="#BA4465",s=15,lw=.85,zorder=5)
    ax.set(xlim=(lo,hi),ylim=(len(chosen)-.45,-1.9))
    labels=LABELS[-2:] if zoom else LABELS
    ax.set_yticks([-1.45,-.85]+list(range(len(chosen))),["二选一","四选一"]+labels)
    ax.set_xticks([50,51,52,53,54] if zoom else [5,10,20,30,40,50,55])
    ax.set_xlabel("r7筛后精判比例 / %",labelpad=3)
    ax.grid(axis="x",alpha=.6);ax.set_axisbelow(True)
    ax.tick_params(labelsize=8)
    if zoom:ax.set_title("(c) K=1000局部放大：r6窗口[50.6416%,51.6812%)",loc="left",fontsize=8.4,pad=7)
    else:ax.set_title(f"({'a' if k==100 else 'b'}) K={k}：整体位置",loc="left",fontsize=9,pad=7)


def figure4(features,intervals,model):
    height=154;fig=fig_mm(height)
    gs=fig.add_gridspec(3,2,height_ratios=[2.2,1.15,.47],left=.14,right=.985,top=.91,bottom=.075,
                          wspace=.35,hspace=.57)
    byname={r["dataset"]:r for r in features};features=[byname[c] for c in CASES]
    feature_panel(fig.add_subplot(gs[0,0]),features,intervals,model,100)
    feature_panel(fig.add_subplot(gs[0,1]),features,intervals,model,1000)
    feature_panel(fig.add_subplot(gs[1,:]),features,intervals,model,1000,zoom=True)
    leg=fig.add_subplot(gs[2,:]);leg.axis("off")
    handles=[Line2D([],[],marker="D",c="#252A2F",ls="",ms=3.5,label="准确比例（事后）"),
             Line2D([],[],marker="o",mfc="white",mec="#168475",ls="",ms=4,label="原始样本比例"),
             Line2D([],[],marker="s",mfc="none",mec="#BA4465",ls="",ms=3.7,label="截断后比例")]
    leg.legend(handles=handles,loc="upper center",ncol=3,frameon=False,fontsize=8,
               handletextpad=.4,columnspacing=1)
    leg.text(.5,-.05,"灰底：校准范围外；色条为冻结选择。准确比例未提供给在线选择器。\n原始越界值保留；K=1不抽样。",ha="center",va="top",fontsize=8)
    fig.suptitle("图4  比例估计与冻结决策区间",y=.982,fontsize=11)
    save(fig,"图4_比例与冻结区间",height,feature_records=6,frozen_interval_records=len(intervals))


def main():
    summary=csv_rows(V/"summary.csv");paired=csv_rows(V/"paired_differences.csv")
    features=csv_rows(V/"sample_feature_errors.csv");intervals=csv_rows(C/"decision_intervals.csv")
    model=json_obj(V/"online_model.json")
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always")
        figure1();figure2(summary);figure3(paired);figure4(features,intervals,model)
    warning_messages=[str(w.message) for w in captured]
    assert not any("Glyph" in m and "missing" in m for m in warning_messages),warning_messages
    assert all(hashlib.sha256((ROOT/path).read_bytes()).hexdigest()==h for path,h in INPUTS.items())
    for record in OUTPUTS:
        svg=OUT/(record["name"]+".svg")
        root=ET.parse(svg).getroot()
        assert not list(root.iter("{http://www.w3.org/2000/svg}image"))
    info={"version":"report-body-figures-v1","status":"generated_from_saved_data_only",
          "width_mm":WIDTH_MM,"png_dpi":300,"font":str(FONT),"font_name":FONT_NAME,
          "matplotlib":matplotlib.__version__,"numpy":np.__version__,"warnings":warning_messages,
          "script_sha256":hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
          "inputs_unchanged":True,"input_sha256":INPUTS,"outputs":OUTPUTS,
          "operations":"CSV/JSON read, exact saved coordinates plotted, linear quantiles recomputed; no random sampling, no query, no fitting, no timing"}
    (OUT/"正文图来源与检查.json").write_text(json.dumps(info,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps({"figures":len(OUTPUTS),"warnings":warning_messages,"width_mm":WIDTH_MM},ensure_ascii=False))


if __name__=="__main__":main()
