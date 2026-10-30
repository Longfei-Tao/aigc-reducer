#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成一份模拟论文 docx，用于试跑与回归测试。

用法：
    python make_sample.py              # 生成 论文样例.docx
    python make_sample.py -o demo.docx

这份样例刻意覆盖了几种「容易改坏」的结构：
    自动编号标题、段内局部加粗、引用标记 [1]、百分比与小数、
    表格、图表题注、参考文献、致谢。
"""

import argparse
import os

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


def build(path):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "宋体"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    def h(text, level=1):
        doc.add_heading(text, level=level)

    def p(text, bold_part=None):
        para = doc.add_paragraph()
        if bold_part and bold_part in text:
            a, b = text.split(bold_part, 1)
            if a:
                para.add_run(a)
            para.add_run(bold_part).bold = True
            if b:
                para.add_run(b)
        else:
            para.add_run(text)
        return para

    h("摘要", 1)
    p("本研究旨在探讨术后早期营养干预对加速康复外科效果的影响。通过对 2023 年 6 月至 2024 年 5 月期间收治的 186 例患者进行回顾性分析，我们认为早期营养干预能够显著提升患者的康复质量，缩短住院时间 [1]。")
    p("研究结果显示，干预组的平均住院天数为 7.2 天，对照组为 10.8 天，两组差异具有统计学意义（P<0.05）[2]。")

    h("1 引言", 1)
    p("加速康复外科（ERAS）理念自提出以来，已经在多个外科领域得到广泛应用。值得注意的是，营养管理作为 ERAS 的核心环节之一，其重要性日益凸显。由此可见，探索一套适合国人的营养干预方案至关重要。")
    p("目前国内相关研究仍然比较少，多数研究样本量偏小，随访时间也比较短，因此结论的推广价值有限。")

    h("2 资料与方法", 1)
    h("2.1 研究对象", 2)
    p("纳入标准包括：年龄 18 至 75 岁；行择期腹部手术；术前营养风险筛查评分 NRS2002 ≥ 3 分。排除标准包括：合并严重心肝肾功能不全者；妊娠期或哺乳期妇女；拒绝签署知情同意书者。共纳入 186 例，其中干预组 93 例，对照组 93 例。")
    h("2.2 干预方案", 2)
    p("干预组于术前 3 天开始口服营养补充剂，每日 2 次，每次 200 毫升；术后 6 小时开始少量饮水，24 小时内启动肠内营养。对照组按常规流程处理。")

    h("3 结果", 1)
    h("3.1 基线资料", 2)
    p("两组患者在年龄、性别、BMI、手术类型等方面的差异均无统计学意义（P>0.05），具有可比性。具体数据见表 1。")

    t = doc.add_table(rows=3, cols=3)
    t.style = "Table Grid"
    rows = [
        ["组别", "例数", "平均年龄（岁）"],
        ["干预组", "93", "52.4±8.6"],
        ["对照组", "93", "53.1±9.2"],
    ]
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            t.cell(i, j).text = cell

    p("表 1 两组患者基线资料比较")
    p("3.2 主要结局指标分析显示，干预组的首次排气时间为 28.6 小时，对照组为 41.3 小时；干预组的术后并发症发生率为 8.6%，对照组为 19.4%，差异具有统计学意义（P<0.01）[3]。综上所论，早期营养干预在加速康复外科中具有较高的临床应用价值。")

    h("4 讨论", 1)
    p("本研究结果与既往多数研究一致。笔者认为，营养干预之所以有效，可能与其维持肠道黏膜屏障功能、减少术后胰岛素抵抗有关。")
    p("当然，本研究也存在一定局限性：单中心回顾性设计、样本量有限、随访时间较短等。未来需要多中心前瞻性研究进一步验证。")

    h("参考文献", 1)
    p("[1] 李明, 张华. 加速康复外科中的营养管理[J]. 中华外科杂志, 2023, 61(4): 289-293.")
    p("[2] 王芳, 刘强. 术后早期肠内营养的临床观察[J]. 肠外与肠内营养, 2022, 29(2): 88-92.")
    p("[3] Gustafsson UO, et al. Guidelines for Perioperative Care in ERAS[J]. World J Surg, 2019, 43(3): 659-695.")

    h("致谢", 1)
    p("感谢导师的悉心指导，感谢科室同事在数据收集过程中给予的帮助。")

    doc.save(path)
    return doc


def main():
    ap = argparse.ArgumentParser(description="生成模拟论文 docx，用于试跑 aigc-reducer")
    ap.add_argument("-o", "--output", default="论文样例.docx")
    args = ap.parse_args()

    out = os.path.abspath(args.output)
    doc = build(out)
    print(f"已生成：{out}")
    print(f"段落 {len(doc.paragraphs)} 段  表格 {len(doc.tables)} 个")
    print("")
    print("下一步（在 examples 目录下执行）：")
    print(f'  python ../skills/aigc-reducer/scripts/docx_tool.py split "{args.output}" -o work/blocks.json')


if __name__ == "__main__":
    main()
