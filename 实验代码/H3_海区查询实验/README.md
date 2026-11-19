# H3 海区查询最小实验

**海区和全部船位都是模拟的。不使用真实海区、真实 AIS，不模拟船舶轨迹。** 本轮只验证固定海区查询的正确性与成本，不设计分辨率选择规则，也不构成完整课程报告。

## 运行

本机已将独立依赖安装在本目录 `.runtime/packages`，未改动 PPT、`PPT_Markdown` 或 MinerU 转换环境。Python 版本要求 3.11+，本机使用 Codex 随附的 Python 3.12。

在本目录打开 PowerShell：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1
```

首次迁移到其他电脑时安装本目录依赖（只写入本目录）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1 -Install
```

运行正确性测试：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1 -TestOnly
```

可以使用 `-PythonPath 'C:\实际路径\python.exe'` 指定 Python。通用 Python 命令如下，代码会自动加载本目录的独立依赖：

```text
python run_experiment.py --seed 20261110 --points 10000 --prepare-repeats 7 --query-repeats 31
```

固定依赖：H3 Python/C 4.5.0，Shapely 2.2.0，NumPy 2.3.5；具体 GEOS、Python、CPU 和系统版本保存到每次运行的 `environment.json`。

## 模拟数据生成规则

- 固定随机种子 `20261110`，NumPy `Generator(PCG64)`。
- 10,000 个随机点，经度在 `[121.875, 122.375)`、纬度在 `[29.875, 30.375)` 分别均匀生成。这是经纬度平面上的均匀分布，不是球面等面积分布。
- 海区有 8 条轴对齐边，从经度 122–122.25、纬度 30–30.25 的矩形顶部挖去一条凹口。顶点见代码中的 `VERTICES` 和每次运行的 GeoJSON。
- 额外加入 32 个测试点：每条边的精确中点及其内、外两侧各一个点（共 24 个），再加 8 个顶点。近邻偏移量是 `1e-7` 度，约厘米量级，专用于检查数值与边界规则，并不代表真实定位精度。
- 总数 10,032 条，记录 ID 从 1 开始且唯一。一个 ID 表示一条观测记录，不进行船舶身份去重。
- 顶点和边中点采用二进制可精确表示的数，边界上的点有独立的人工预期答案。海区边界统一计入，不对海区使用模糊缓冲。

本轮的精确几何语义是小范围经纬度平面上的直线多边形。H3 格子轮廓由 `cell_to_boundary` 返回的顶点连接而成，在同一平面内做相交与包含判定。不讨论球面测地面积、跨日期线或极区问题。

## 公平对照与覆盖方式

**直接方法**：先用 NumPy 做外包矩形筛选（含边界），再用经过 `prepare` 预处理的 Shapely `intersects_xy` 批量精判。它避免 Python 逐点循环，也避免为每个点创建几何对象。对于一个点，`intersects_xy(海区, x, y)` 包含边界，与 `covers(海区, Point)` 的查询语义一致。

**H3 r6/r7/r8**：每档都从原始纬度、经度独立调用 `latlng_to_cell`。将编码结果排序、分组，构造“格子编号 → 原始记录位置数组”的索引。每次查询按内部格和边界格查索引，内部点直接纳入，边界点用同一个预处理海区和 `intersects_xy` 精判，最后返回 ID 数组。不缓存固定海区对应的候选记录数组或最终答案。

**完整覆盖**：使用官方实验接口 `h3shape_to_cells_experimental(..., contain='bbox_overlap')` 保守枚举，随后将返回的格子轮廓与海区做 `intersects`，删除不相交的外包候选。`intersects` 包含仅接触边界的情况。对保留的相交格，用 `covers(海区, 格子轮廓)` 判定内部格，其余是边界格。没有采用仅按格子中心落入的覆盖方式。

覆盖审计在计时之外执行：

1. 检查覆盖格的几何并集是否完整覆盖海区，保存未覆盖面积。
2. 从比海区外包矩形向外再扩展 `0.0625` 度的大矩形重新枚举，检查其中与海区相交的格子集合是否完全相同。
3. 核对最终记录 ID 与直接精判参照集合一致，单独检查 32 个边界点。

这是对本次具体几何与数据的审计，不将测试通过解释为通用球面算法证明。实验接口版本固定到 H3 4.5.0。

## 计时范围

数据生成、依赖导入、文件读写、集合比较、正确性检查、覆盖审计都不在核心计时内。计时单位为原始纳秒，结果汇总转为毫秒。

准备阶段每方案重复 **7 次**，每次重建，分别记录：

1. 海区几何创建、预处理及触发延迟准备的时间。
2. 点索引总时间，另拆出 H3 编码与排序分组时间。直接方法不需要点索引。
3. 保守覆盖格枚举时间。
4. 格子轮廓构造、相交过滤、内部/边界分类时间。
5. 准备总时间及准备后的第一次查询时间；将同一试次二者相加，得到“只查一次”的成本。

准备后查询每方案重复 **31 次**，共用相同输入与同类输出契约：返回唯一记录 ID 的 NumPy 数组，不要求记录顺序。每次都包含候选查找、坐标提取、精判和结果数组构造，另记录三个分项。集合比较在查询计时结束后进行，未为任一方案强加排序要求。

所有方案采用同样的预热策略，每轮随机打乱方案顺序，计时期间统一关闭 Python 循环垃圾回收，结束后恢复。重复查询是内存中数据、索引不变的运行状态，不是含导入和磁盘读取的全进程冷启动。重复查询不直接复用上一次答案。数据或海区更新需要重建，本轮不模拟更新。

每列报告自己的中位数。一次性总成本先在每个试次里将准备与首查相加，再对总数取中位数；不将分项中位数硬相加。查询时间不足 1 毫秒，细小差异需要结合每次原始记录解释。

## 输出

每次运行创建新的 `runs/日期_时间/`，旧运行不会覆盖。根目录 `LATEST_RUN.txt` 指向最新成功运行，`结果速览.md` 是简短结果表。

- `data/`：模拟点 CSV、海区 GeoJSON、完整生成规则。
- `point_indexes/`：各档每条原始记录的 H3 编码。
- `coverage/`：全部覆盖格编号及内部/边界分类。
- `selected_ids/`：各方案最终选中的记录 ID，供独立复核。
- `correctness.json`：误选、漏选、重复 ID、覆盖审计及点数统计。
- `boundary_cases.csv`：32 个测试点的预期、参照答案、各方案答案与所在格分类。
- `diagnostics/incorrect_points.csv`：有差异时先写入出错点，再中止性能实验；无差异时只保留列名。
- `preparation_trials.csv`：每次准备、分项、首查与一次性总成本的原始计时。
- `query_trials.csv`：每次重复查询的原始计时与候选/精判点数。
- `summary.csv`、`summary.json`：所有结果的中位数、最小/最大查询耗时与正确性。
- `environment.json`、`parameters.json`、源码与依赖快照、`sha256_manifest.json`：复现环境、参数与内容指纹。

## 官方依据

- [H3 实验填格接口及覆盖模式](https://uber.github.io/h3-py/api_verbose.html#h3.h3shape_to_cells_experimental)
- [Shapely intersects_xy：批量点坐标判断](https://shapely.readthedocs.io/en/stable/reference/shapely.intersects_xy.html)
- [Shapely prepare：几何预处理](https://shapely.readthedocs.io/en/stable/reference/shapely.prepare.html)
- [Shapely covers：包含边界的覆盖语义](https://shapely.readthedocs.io/en/stable/reference/shapely.covers.html)
