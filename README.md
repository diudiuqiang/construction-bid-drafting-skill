# construction-bid-drafting-skill

电力 EPC 技术标「施工方案」章节的编写、合并编号、Word 图表混排生成与合规 / 视觉自检 skill（WorkBuddy）。

## 能力

- **编制模式**：分册并行撰写 → 合并编号 → 生成 docx（大纲层级、元数据去工具痕迹）
- **体检模式**：禁用词 / 硬指标 / 编号 / 结构审计 / AI 痕迹 五类体检
- **合卷交付模式**：跨章号重映射（如 3~8 → 10~13）、Excel 拆表入 Word、内嵌图片与图表题注
- **视觉终检**：LibreOffice 转 PDF + pymupdf 核对图片渲染、题注居中、字体 / 字号 / 越界
- **二次修正机制**：设计 / 勘察资料到手后增量修订而不重写

## 结构

```
SKILL.md
references/   写作基准、合规红线、章号重构与合卷、Word 图表混排与视觉终检、校验器误判边界 等
scripts/      build_docx / merge_validate / compliance_check / ai_trace / audit_structure / cbd_common
assets/       待修正清单模板
```

## 使用方法

将本目录放入 `~/.workbuddy/skills/` 即可被 WorkBuddy 识别加载。

> 本 skill 的部分硬指标、禁用词、工程量口径为项目级配置，应放在各自的 `conf/` 中，勿写死进通用 skill。
