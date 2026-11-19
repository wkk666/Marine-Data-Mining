"""只整理已完成验证记录，不生成数据、不运行查询、不拟合或改变决策。"""
import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream))


def main():
    run = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / (ROOT / "LATEST_INDEPENDENT_VALIDATION.txt").read_text(encoding="utf-8").strip()
    comparisons = read_csv(run / "comparisons.csv")
    paired = read_csv(run / "paired_differences.csv")
    pairs = [("sample_two", "k_only"), ("sample_four", "sample_two"), ("sample_four", "fixed_direct"), ("sample_four", "fixed_r7")]
    lookup = {(r["dataset"], r["K"], r["strategy"], r["control"]): r for r in comparisons}
    lines = [
        "# 本轮结果要点与净节省表", "",
        "这是已完成独立验证的结果整理，不是完整实验报告。仅从已保存的记录计算，没有运行新查询、生成新数据或修改模型。", "",
        "正确性通过：6组模拟数据、18个场景、7种执行方案、每组合7次，正式整体计时共882条；误选、漏选、重复记录均为0，原32个边界点及覆盖审计通过。54条推荐先锁定，随后才做预检、预热和计时；正式任务重做选择并与推荐核对一致。", "",
        "每组为10000个随机点加原32个边界测试点。中等/较强聚集的40%/80%是边界带生成成分比例，背景均匀点仍可能落在边界附近，不能将其当作实际边界附近点占比。海区、0.002度带宽与拒绝采样方法均沿用原实现。", "",
        "## 应当保留的结论", "",
        "1. 本轮没有显示200点抽样相对K-only具有稳定净收益。K=100时仅uniform_3102的样本二选一小幅节省0.0910 ms（0.545%），另一个均匀种子反而亏损；K=1000时六组全部更慢，按总时间中位数比较多花1.3848—3.2943 ms。K=1均不抽样且都选direct，微小正负差没有方法选择上的收益。", "",
        "2. 样本二选一和四选一在全部18个场景选择相同。两者的正负时间差反映执行上下文和计时波动，不能算新增分辨率收益。**本验证集未触发新增档位，未验证其被选中后的收益。**", "",
        "3. 完整四选一相对始终直接查询：K=1000的两组均匀数据净省74.1263/77.4311 ms（46.03%/46.30%），两组中等聚集净省32.2304/34.4999 ms（12.74%/13.24%）；强聚集反而多花1.8359/2.0852 ms。K=100的六组均无净收益。", "",
        "4. 完整四选一相对始终r7：K=1通过直接查询省6.65—6.81 ms；K=100的中等/强聚集省0.8046—4.4307 ms（2.83%—10.88%），但均匀两组更慢；K=1000六组全部更慢。选择的价值取决于对照，不能只选有利的对照描述。", "",
        "样本策略选中的固定方法与事后最优一致16/18场景，K-only为14/18；这个命中率改善没有自动变成净省时。18场景仍只有6组独立生成数据；K及同一种子下的7次计时不增加独立数据数量。", "",
        "## 为什么有些场景仍然更慢", "",
        "抽样误差确实存在。moderate_3202的准确比例为0.308513，样本比例为0.365，偏高5.649个百分点；其K=1000的r7固定方法成本预测误差，使用准确比例为−1.18%，使用样本比例变成+13.80%。预测误差依据正常固定方法整体计时，不含选择费；预测函数仍按冻结规则将比例截断到校准范围。", "",
        "不过本轮用准确比例替换样本比例的事后反事实诊断没有改变任何样本策略推荐。这只说明本轮偏差未越过决策分界，不能据此说200点足够准确或对其他数据无风险。", "",
        "K=100均匀数据中，抽样使策略从K-only的direct改为r7，固定方法确实更快，但收益约1.6—1.9 ms，通常不足以覆盖选择流程增量。K=100中等/强聚集，以及K=1000均匀/中等聚集，两种策略往往选择相同方法，额外抽样没有带来决策收益。", "",
        "两组强聚集、K=1000的固定方法中位数依次为：strong_3301的r6/r7/direct = 346.5219/347.3882/348.7875 ms；strong_3302 = 349.8512/351.2331/352.6519 ms。r6确实是本轮事后最快档，但相对r7仅省0.8663/1.3819 ms（0.25%/0.39%），相对direct仅省0.65%/0.79%。这些是观察到的差额，不能据7次计时宣称统计显著。", "",
        "强聚集准确比例0.518740/0.523325，均超过冻结r6入选区间的上界0.516812。原始四方法模型最低成本均为r6；加入保守规则后，即使用准确比例仍选direct。这里体现保守规则放弃的小收益及选择流程增量，不能主要归咎于抽样改选或原始模型排序错误。2%是事先固定的工程余量，不是实测差异的显著性门槛，本轮不据此调整。", "",
        "策略与其选中固定方法的中位数差包含选择过程、执行上下文和计时波动，不能把这整个差额精确命名为独立抽样费用；正式计时没有加入分项探针。", "",
        "## 全部18场景的净节省", "",
        "净节省按‘对照实测总时间中位数减策略实测总时间中位数’计算。每格为毫秒（占对照的百分比）；负数表示策略更慢。此表是两个中位数之差，不是逐轮差值的中位数。", "",
        "| 数据 | K | 二选一相对K-only | 四选一相对二选一 | 四选一相对直接 | 四选一相对固定r7 |",
        "| --- | ---: | ---: | ---: | ---: | ---: |",
    ]
    cases = sorted({(r["dataset"], int(r["K"])) for r in comparisons}, key=lambda x: (int(x[0].split("_")[-1]), x[1]))
    for name, k in cases:
        values = []
        for strategy, control in pairs:
            r = lookup[name, str(k), strategy, control]
            values.append(f"{float(r['median_difference_saving_ms']):+.4f}（{float(r['median_difference_saving_percent']):+.2f}%）")
        lines.append(f"| {name} | {k} | " + " | ".join(values) + " |")
    lines += ["", "## 最大损失：区分场景中位数与某一轮", "",
        "场景损失取两个总时间中位数之差的相反数；单轮损失取同一计时轮次中策略总时间减对照总时间。毫秒最坏值与百分比最坏值分别在各自指标上取最大，未假定来自同一轮；括号保留同一轮的另一种单位，原始逐轮值在paired_differences.csv。", "",
        "| 策略相对对照 | 最大场景中位数损失 ms（%） | 最大单轮额外耗时 ms（该轮%） | 最大单轮额外比例 %（该轮ms） | 场景 |",
        "| --- | ---: | ---: | ---: | --- |"]
    for strategy, control in pairs:
        rows = [r for r in comparisons if r["strategy"] == strategy and r["control"] == control]
        median_worst = min(rows, key=lambda r: float(r["median_difference_saving_ms"]))
        raw_pairs = [r for r in paired if r["strategy"] == strategy and r["control"] == control]
        paired_worst = min(raw_pairs, key=lambda r: float(r["net_saving_ms"]))
        percent_worst = min(raw_pairs, key=lambda r: float(r["net_saving_percent"]))
        lines.append(f"| {strategy} 相对 {control} | {-float(median_worst['median_difference_saving_ms']):.4f}（{-float(median_worst['median_difference_saving_percent']):.3f}%） | {-float(paired_worst['net_saving_ms']):.4f}（{-float(paired_worst['net_saving_percent']):.3f}%） | {-float(percent_worst['net_saving_percent']):.3f}%（{-float(percent_worst['net_saving_ms']):.4f}） | 中位数最坏：{median_worst['dataset']}/K{median_worst['K']}；单轮毫秒最坏：{paired_worst['dataset']}/K{paired_worst['K']}/轮{paired_worst['trial']}；单轮比例最坏：{percent_worst['dataset']}/K{percent_worst['K']}/轮{percent_worst['trial']} |")
    lines += ["", "例如strong_3302、K=1000：样本二选一相对K-only按两个中位数之差多花3.2943 ms；逐轮净节省中位数为−4.5294 ms，第25—75百分位为−6.6659至−2.7004 ms，最坏一轮多花24.4056 ms。三者是不同统计量。", "",
        "四选一在该场景相对事后固定最优r6多花4.8859 ms（约1.40%），其中固定方法选择差额2.8007 ms，其余2.0852 ms是策略流程与计时上下文的合计差额。事后参考取四个固定方法总时间中位数的最小值，没有挑原始计时中偶然最快的一次。", "",
        "## 保存材料与停止点", "",
        "独立验证结果.md含126组总时间中位数及第25—75百分位；summary.csv含最小/最大等统计；formal_trials.csv保留882条原始整体计时；comparisons.csv及paired_differences.csv分别保留两类差值统计。预先推荐、样本位置、模型成本、正确性、事后诊断、原始数据、代码模型协议快照与SHA256均在本目录。", "",
        "原模型、样本量、抽样种子、2%门槛、数据比例、带宽均未调整。没有追加种子、重新校准、完整报告或GitHub发布。验证到此停止。", ""]
    lines += ["## 只读复核与指纹", "",
        "artifact_validation.json保存事后产物审计结果。原sha256_manifest.json保持不变；delivery_sha256_manifest.json另行覆盖新增的审计、解读文件及事后整理脚本。postprocessing/中的脚本在正式实验结束后新增，没有冒充生成数据前已冻结的代码。", "",
        "下列命令仅复核已保存记录，不生成数据、不运行查询或性能实验：", "", "```powershell",
        "& 'C:\\Users\\wkc\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe' 'D:\\CodeX\\海洋数据挖掘\\H3_海区查询实验\\verify_independent_validation.py'", "```", ""]
    (run / "结果要点与净节省表.md").write_text("\n".join(lines), encoding="utf-8")
    print(run / "结果要点与净节省表.md")


if __name__ == "__main__":
    main()
