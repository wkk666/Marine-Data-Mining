# 实验代码与复现

完整项目位于 [H3_海区查询实验](H3_海区查询实验/)。阶段入口共用该项目中的查询、选择与数据生成代码。先阅读[项目总览与复现](H3_海区查询实验/项目总览与复现.md)。

```powershell
cd H3_海区查询实验
python verify_independent_validation.py
```

从本目录执行上面的命令。核对入口只读，使用 Python 标准库，检查保存的结果、统计和指纹链。执行下表的重算命令前，在项目根目录用 Python 3.12 环境安装 [requirements.txt](H3_海区查询实验/requirements.txt) 中的实验依赖。

| 阶段 | 入口 | 重算命令 | 当时的记录 |
| --- | --- | --- | --- |
| 最初小验证 | [`run_experiment.py`](H3_海区查询实验/run_experiment.py) | `python run_experiment.py` | [已保存运行](H3_海区查询实验/runs/20261110_205508_135429/) |
| 四组数据探索 | [`exploration_experiment.py`](H3_海区查询实验/exploration_experiment.py) | `python exploration_experiment.py` | [已保存运行](H3_海区查询实验/runs/explore_20261110_221236_751069/) |
| 矩形筛选对照 | [`bbox_ablation.py`](H3_海区查询实验/bbox_ablation.py) | `python bbox_ablation.py` | [已保存运行](H3_海区查询实验/runs/bbox_20261111_001350_991724/) |
| 选择器校准 | [`calibrate_selector.py`](H3_海区查询实验/calibrate_selector.py) | `python calibrate_selector.py` | [已保存运行](H3_海区查询实验/runs/selector_cal_20261111_140827_086716/) |
| 六组独立验证 | [`run_independent_validation.py`](H3_海区查询实验/run_independent_validation.py) | `python run_independent_validation.py` | [已保存运行](H3_海区查询实验/runs/validation_20261111_151150_257869/) |

重算会生成新的运行目录，并可能更新对应的 `LATEST_*.txt`。依照[独立验证复现说明](H3_海区查询实验/独立验证复现说明.md)使用冻结参数与种子，耗时会受本机环境影响。报告引用已保存的 `validation_20261111_151150_257869`。

统一准备和执行见 [selector_engine.py](H3_海区查询实验/selector_engine.py)，冻结预测与决策见 [selector_decision.py](H3_海区查询实验/selector_decision.py)，直接查询与 H3 覆盖见 [run_experiment.py](H3_海区查询实验/run_experiment.py)，新版边界筛选见 [bbox_ablation.py](H3_海区查询实验/bbox_ablation.py)。

正文图稿可直接查看，重画环境见[绘图依赖](H3_海区查询实验/报告材料/requirements_绘图.txt)和[总览](H3_海区查询实验/项目总览与复现.md)。绘图脚本依赖 Windows 微软雅黑字体，部分旧报告辅助命令绑定原制作机器。
