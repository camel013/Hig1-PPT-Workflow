# Hig1 PPT Workflow

本仓库保存 PPT 任务的执行纪律、生成脚本、设计规范、原始素材和最终可交付成果。大型二进制文件使用 Git LFS，生成预览、草稿、备份和重复检查图不入库。

## 仓库内容

- `AGENTS.md`：版式设计、读取边界、执行预算与验收停止条件。
- `_ppt_workflow/`：工作区预处理与检查工具。
- `projects/`：各项目的设计规范、生成脚本、SVG、原始素材和最终 PPTX。
- `glass_layer_sample_20261026/`：玻璃质感设计体系样页。
- `TEMPLATE_LOCK.json`、`TEMPLATE_LOCK.sha256`：固定模板与已完成项目的锁定索引。
- `template_lock.ps1`：模板锁定校验脚本。

## 不入库内容

- 各类页面预览、联系表、裁切图、临时构建目录和检查图。
- `_src/`、`_work/`、`backups/`、重复统一版和临时质量基线。
- 运行日志、缓存和草稿构建物。

## 获取仓库

只需要读取规范和脚本时，使用文本浅克隆，不下载 LFS 资产：

```powershell
$env:GIT_LFS_SKIP_SMUDGE = "1"
git clone https://github.com/camel013/Hig1-PPT-Workflow.git
```

需要某个项目的素材时，只拉取对应目录：

```powershell
git lfs pull --include="projects/campus_environment_ppt169_20261025/images/**"
```

需要完整复现时再执行 `git lfs pull`。

## 使用原则

1. 新任务先读取相关单文件，不扫描整个项目目录。
2. 复用模板和生成脚本，局部修改后只检查变更页。
3. 字体嵌入前仅做结构和字体名称检查；嵌入后对最终产物完成一次全页验收。
4. 超过 8 页先生成联系表，只放大异常页面。
5. 验收通过立即停止；同一失败修复两次仍未收敛时报告并退出。
