# 验收记录

长稿任务在本地研究档案保存简明回执；短文本可以在说明中交代。记录真实情况，不预填通过：

```json
{
  "task_scope": "本次修改范围",
  "input_sha256": "原稿指纹",
  "profile_id": null,
  "language_review": "completed | partial | not_checked",
  "claim_and_citation_check": "completed | partial | not_checked",
  "section_review": {},
  "citation_support_and_placement": "not_checked",
  "citation_numbering": "not_checked",
  "bibliographic_metadata": "not_checked",
  "bibliographic_style": "not_checked",
  "ooxml_check": "completed | partial | not_checked",
  "field_update": "completed | partial | not_checked",
  "visual_review": {"status": "not_checked", "pages_reviewed": []},
  "author_readback": "not_attested",
  "institutional_acceptance": "not_attested",
  "open_items": []
}
```

每个 completed 应有范围与证据路径；只检查部分段落或部分页就记录 partial。完成机器检查也不能将作者/导师/学院状态改为通过。缺少字体、真实字段更新或渲染器时明确下一项检查，不用“已生成Word”代替排版验收。

全稿的 `section_review` 分别登记目录、中英文摘要、绪论、文献理论、方法、现状诊断、方案实施、评价讨论、结论、文内引用和文末文献；合并章节仍按实际任务登记，缺少或不适用写明原因。引用的支持关系、位置、编号、题录和著录格式分别核查，不能用一次正则扫描替代全部项目。

发表或共享前检查文本、修订人、批注、文档属性、图片、嵌入对象和附录的个人/商业信息。开源技能的许可不覆盖输入材料。提交、签字、授权、学校评阅等状态只依真实主体行为登记。
