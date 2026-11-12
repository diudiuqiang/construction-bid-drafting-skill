# construction-bid-drafting

电力 EPC 项目「施工方案 / 施工组织设计」章节的一体化编写 skill：分册并行编写 → 合并编号 / 章号重构合卷 → Word 图表混排生成 → 合规与视觉自检。

## 能力概览

- **分册并行编写 + 两遍法合并编号**：自动重排表/图号，跨册引用感知，杜绝悬空引用与整体号段偏移。
- **Word 图表混排生成**（OOXML）：支持 `[[IMG:相对路径]]` 图片插入、题注居中判定、CJK 列宽自适应、LibreOffice + PyMuPDF 视觉终检。
- **合规体检**：禁用词 / 硬指标 / 章节节点 / 编号 / 初稿注五类检查；配套 AI 痕迹识别消解与结构审计。

## 目录结构

| 目录 | 内容 |
|---|---|
| `SKILL.md` | 主入口，三模式：编制 / 体检 / 合卷交付 |
| `references/` | 11 篇（写作基准、章节骨架、合规红线、二次修正流程、Word 图表混排与视觉终检、校验器误判边界等） |
| `scripts/` | 6 个（`build_docx` / `merge_validate` / `compliance_check` / `ai_trace` / `audit_structure` / `cbd_common`） |
| `assets/` | 待修正清单模板 |

## ⚡ 本仓库的发布方式（重点）

> **本仓库不使用 `git push` 发布，而是由本地 WorkBuddy 的 `github-skill-sync` skill 内置脚本 `publish_skill.py`（新脚本）经 `api.github.com` 反向同步发布。**

这样做的原因与本脚本的工作机制：

- **网络约束**：本机 `github.com:443` 被墙，`git push` 不可用；唯一可达通道是 `api.github.com` REST API。
- **发布流程**：脚本走「`GET /user` 验证 token → 幂等探测仓库（404 才建库，不删库重建）→ 逐文件 `PUT`：新文件不带 `sha`、已存在先 `GET sha` 再带 `sha` 防 `409` → `git/trees` 验证远端文件数」的标准 Contents API 链路。
- **本次提交即由该【新脚本】完成**：所有 `.md` / `.py` 的中文路径与正文经 base64 编码上传，无乱码；增量更新时仅对发生变化（sha 不同）的文件重新上传。

如需自行发布或增量更新本 skill 到 GitHub，在本机执行：

```bash
set GITHUB_TOKEN=你的经典PAT(需 repo 权限，用完建议撤销)
python publish_skill.py --skill-dir C:/Users/youqw/.workbuddy/skills/construction-bid-drafting --repo construction-bid-drafting-skill
```

> 脚本位于 `github-skill-sync` skill 的 `scripts/` 目录；支持 `--private` 建私有库、`--dry-run` 只列文件不实际请求。

## 使用方式

在 WorkBuddy 中调用 `construction-bid-drafting` skill，按 `SKILL.md` 的三模式推进编制 / 体检 / 合卷交付。项目专属口径（如鱼塘分块、峰值人数、机械台套、禁用词清单）一律由项目 `conf/` 承载，不写死进本通用 skill。
