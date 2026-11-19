"""仅拼接已经写好的正文/附录，不运行实验、选择或拟合。"""
import hashlib
import json
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPORT = HERE / "实验报告_完整初稿.md"
MARKER = "<!-- APPENDICES_FROM_SAVED_RECORDS -->"


def main():
    body = REPORT.read_text(encoding="utf-8").split(MARKER)[0].rstrip()
    appendix = (HERE / "完整报告_附录数据片段.md").read_text(encoding="utf-8")
    appendix = appendix.replace(
        "本片段表格仅由既存CSV/JSON读取、格式化生成；未重新生成数据、运行查询、计时、拟合或决策。报告引用路径均从“报告材料”目录出发。",
        "附录数值读取既有CSV/JSON，没有重新测量或拟合。下列文件链接均相对于本报告所在目录。")
    appendix = appendix.replace("口径A", "中位数差口径").replace("口径B", "同轮差中位数口径")
    appendix = appendix.replace("r6的入选窗口只在K=1000上述窄区间存在", "在所分析的K=100、1000中，r6的入选窗口只在K=1000上述窄区间存在")
    index = (HERE / "完整报告_材料索引补充.md").read_text(encoding="utf-8")
    complete = body + "\n\n" + MARKER + "\n\n" + appendix.rstrip() + "\n\n" + index.rstrip() + "\n"
    complete = re.sub(r"(?<!较)强聚集", "较强聚集", complete)
    REPORT.write_text(complete, encoding="utf-8")
    print(json.dumps({"file":REPORT.name,"characters":len(REPORT.read_text(encoding='utf-8')),
                      "sha256":hashlib.sha256(REPORT.read_bytes()).hexdigest()},ensure_ascii=False))


if __name__ == "__main__":
    main()
