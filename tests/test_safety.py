"""危险输入的回归：拒绝损坏文档、错位回写及覆盖输入。"""
import json
import unittest

from test_pipeline import Base, Document, OxmlElement, docx_tool, make_doc, run


class TestSafety(Base):
    def edits(self, name, data):
        path = self.dir / name
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return path

    def test_plain_reference_heading_is_protected(self):
        doc = Document()
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 示例文献。")
        self.assertNotEqual(docx_tool.scan_blocks(doc)[0]["kind"], "para")

    def test_body_mention_of_references_does_not_end_document(self):
        doc = Document()
        doc.add_paragraph("参见参考文献。")
        doc.add_paragraph("这里仍然是正文。")
        self.assertEqual([b["kind"] for b in docx_tool.scan_blocks(doc)], ["para", "para"])

    def test_mixed_image_and_field_paragraphs_are_protected(self):
        for tag in ("w:drawing", "w:fldChar", "w:object", "w:instrText"):
            with self.subTest(tag=tag):
                doc = Document()
                p = doc.add_paragraph("这里是正文和不可破坏的对象。")
                p.runs[0]._r.append(OxmlElement(tag))
                self.assertNotEqual(docx_tool.scan_blocks(doc)[0]["kind"], "para")

    def test_ambiguous_section_is_rejected(self):
        doc = self.dir / "ambiguous.docx"
        make_doc(doc, ["研究方法一", "正文甲。", "研究方法二", "正文乙。"], heading_at=(0, 2))
        blocks = self.dir / "ambiguous.json"
        run("split", doc, "-o", blocks)
        code, _, _ = run("batch", blocks, "--section", "研究方法")
        self.assertNotEqual(code, 0)

    def test_check_failure_has_nonzero_exit(self):
        edits = self.edits("lost.json", {1: "遗漏所有引用和数字。"})
        code, _, _ = run("check", self.sample, edits, "-o", self.dir / "lost-check.json")
        self.assertEqual(code, 1)

    def test_citation_occurrence_loss_is_detected(self):
        src = self.dir / "cites.docx"
        make_doc(src, ["研究 [1] 支持甲，另一研究 [1] 支持乙。"])
        edits = self.edits("cites.json", {0: "研究 [1] 支持甲，也支持乙。"})
        report = self.dir / "cites-check.json"
        run("check", src, edits, "-o", report)
        self.assertEqual(json.loads(report.read_text(encoding="utf-8"))["verdict"], "fail")

    def test_added_citation_and_changed_percentage_fail(self):
        src = self.dir / "numbers.docx"
        make_doc(src, ["结果为 30%，原研究 [1] 支持这一观察。"])
        edits = self.edits("numbers.json", {0: "结果为 30，研究 [1][2] 支持这一观察。"})
        report = self.dir / "numbers-check.json"
        run("check", src, edits, "-o", report)
        self.assertEqual(json.loads(report.read_text(encoding="utf-8"))["verdict"], "fail")

    def test_auxiliary_output_cannot_overwrite_source(self):
        src = self.dir / "aux-source.docx"
        src.write_bytes(self.sample.read_bytes())
        edits = self.edits("alias.json", {1: "改写正文。"})
        before = src.read_bytes()
        code, _, _ = run("write", src, edits, "-o", self.dir / "alias-out.docx", "--diff-md", src)
        self.assertNotEqual(code, 0)
        self.assertEqual(src.read_bytes(), before)

    def test_existing_output_is_not_overwritten(self):
        edits = self.edits("existing.json", {1: "改写正文。"})
        out = self.dir / "existing.docx"
        out.write_bytes(b"keep")
        code, _, _ = run("write", self.sample, edits, "-o", out)
        self.assertNotEqual(code, 0)
        self.assertEqual(out.read_bytes(), b"keep")

    def test_source_changed_since_split_is_rejected(self):
        src = self.dir / "drift.docx"
        make_doc(src, ["原来的正文。"])
        blocks = self.dir / "drift-blocks.json"
        run("split", src, "-o", blocks)
        edits = self.edits("drift-edits.json", {0: "改写后的正文。"})
        code, _, err = run("write", src, edits, "--blocks", blocks, "-o", self.dir / "valid.docx")
        self.assertEqual(code, 0, err)
        make_doc(src, ["插入的新正文。", "原来的正文。"])
        out = self.dir / "drift-out.docx"
        code, _, _ = run("write", src, edits, "--blocks", blocks, "-o", out)
        self.assertNotEqual(code, 0)
        self.assertFalse(out.exists())

    def test_duplicate_ids_are_rejected(self):
        code, _, _ = run("apply", self.blocks, "--ids", "1,1", "--dry-run", stdin="甲。\n\n乙。".encode("utf-8"))
        self.assertNotEqual(code, 0)

    def test_comparison_direction_change_is_detected(self):
        src = self.dir / "comparison.docx"
        make_doc(src, ["差异有统计学意义，P<0.05。"])
        edits = self.edits("comparison.json", {0: "差异有统计学意义，P>0.05。"})
        code, _, _ = run("check", src, edits, "-o", self.dir / "comparison-check.json")
        self.assertEqual(code, 1)

    def test_strip_number_preserves_decimal(self):
        self.assertEqual(docx_tool.split_reply("2023.5 是原始数值。", True), ["2023.5 是原始数值。"])

    def test_mixed_note_run_is_protected(self):
        doc = Document()
        p = doc.add_paragraph("正文与脚注在同一个 run。")
        p.runs[0]._r.append(OxmlElement("w:footnoteReference"))
        self.assertEqual(docx_tool.scan_blocks(doc)[0]["kind"], "complex")

    def test_apply_edits_are_bound_to_source_without_blocks_flag(self):
        edits = self.dir / "bound-edits.json"
        code, _, err = run("apply", self.blocks, "--batch", 1, "-o", edits,
                           stdin="甲。\n\n乙。".encode("utf-8"))
        self.assertEqual(code, 0, err)
        src = self.dir / "other-source.docx"
        make_doc(src, ["完全不同的正文。", "同 id 的正文。", "下一段正文。"])
        out = self.dir / "wrong-source.docx"
        code, _, _ = run("write", src, edits, "-o", out)
        self.assertEqual(code, 1)
        self.assertFalse(out.exists())

    def test_apply_cannot_overwrite_indirect_source(self):
        src = self.dir / "indirect.docx"
        make_doc(src, ["原稿正文。"])
        blocks = self.dir / "indirect-blocks.json"
        run("split", src, "-o", blocks)
        before = src.read_bytes()
        code, _, _ = run("apply", blocks, "--batch", 1, "-o", self.dir / "indirect-edits.json", "--progress", src,
                         stdin="新正文。".encode("utf-8"))
        self.assertEqual(code, 1)
        self.assertEqual(src.read_bytes(), before)

    def test_end_to_end_reopens_generated_document(self):
        src = self.dir / "flow.docx"
        make_doc(src, ["研究纳入 30 例[1]。其中有效 20 例。"])
        before = src.read_bytes()
        blocks = self.dir / "flow-blocks.json"
        edits = self.dir / "flow-edits.json"
        report = self.dir / "flow-check.json"
        out = self.dir / "flow-out.docx"
        reply = "研究纳入 30 例[1]，其中有效 20 例。"
        for command in (
            ("split", src, "-o", blocks),
            ("batch", blocks, "--id", 1, "-o", self.dir / "flow-batch.txt"),
        ):
            code, _, err = run(*command)
            self.assertEqual(code, 0, err)
        code, _, err = run("apply", blocks, "--batch", 1, "-o", edits, stdin=reply.encode("utf-8"))
        self.assertEqual(code, 0, err)
        for command in (
            ("check", src, edits, "--blocks", blocks, "-o", report),
            ("write", src, edits, "--blocks", blocks, "-o", out),
            ("status", blocks, "-e", edits),
        ):
            code, _, err = run(*command)
            self.assertEqual(code, 0, err)
        self.assertEqual(json.loads(report.read_text(encoding="utf-8"))["verdict"], "pass")
        self.assertEqual(Document(str(out)).paragraphs[0].text, reply)
        self.assertEqual(src.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
