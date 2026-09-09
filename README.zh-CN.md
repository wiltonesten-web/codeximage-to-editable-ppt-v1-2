# codeximage-to-editable-ppt-v1-2

**简体中文** | [English](README.md)

这是 [`codeximage-to-editable-ppt-v1`](https://github.com/wiltonesten-web/codeximage-to-editable-ppt-v1)
的升级编排层。它面向多页图片型 PowerPoint，在完整保留 V1 单页质量标准的同时，
通过经过测量的并行执行完成可编辑化重建。可以并行处理 3、6、9 等数量的图片；
实际并行规模取决于可用的任务树和硬件容量。

> **效率提升：** 通过独立页面 Worker 同时重建多张幻灯片，在不降低 V1 单页质量标准的前提下，
> 减少多页任务的总等待时间。实际效率提升以运行中测得的 Worker 时间重叠和耗时结果为准。

当前打包的 Skill 版本为 **1.2.7**。该版本新增硬件感知预检、多根任务编排、
可复用品牌页眉、更严格的科学图表与分图标记检查、结构化信息卡拆分规则、
有序 OpenXML 合并以及合并后的内容不变量检查。

## 主要特性

- 在向正式输出目录写入文件之前，对多页幻灯片图片或图片型 PPT/PPTX 页面进行预检。
- 根据 CPU、可用内存、正在运行的 PowerPoint 进程和可用 Codex Worker 槽位，保守估算并发数量。
- 每个 Codex 任务树运行最多 3 个页面 Worker；在容量允许且用户明确授权创建相应任务树时，
  可以扩展到 3、6、9 等页数。
- 检测重复出现的品牌页眉，并在统一的归一化位置复用一份经过哈希验证、不可变的裁剪图。
- 拒绝复用整页截图、可编辑文字下方残留栅格文字、信息卡拆分不足以及页眉与标题发生碰撞等问题。
- 将每个逻辑科学图表保留为一张与原图像素完全一致的裁剪图，并包含标题、坐标轴、刻度、数值、单位、图例、注释、色标和安全边距。
- 保留 `(a)`、`(b)`、`(c)` 等外置科学分图标记，并对标记与文字进行精确校验。
- 按源页面顺序合并已通过检查的单页演示文稿，保留审计标签，并快速验证合并后的内容不变量。
- 报告实际观测到的 Worker 时间重叠，不会宣称未经测量的并行度。

## 与 V1 的关系

本仓库是 V1 的升级层，而不是 V1 的替代副本。安装后的 V1.2 Skill 会从以下同级目录读取
V1 的标准单页重建规范：

```text
%USERPROFILE%\.codex\skills\codeximage-to-editable-ppt-v1
```

请先安装 V1，再安装本仓库中的 V1.2 Skill。两个 Skill 文件夹必须并列保留。

## 安装为 Codex Skill

安装所需的 V1 基础 Skill：

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo wiltonesten-web/codeximage-to-editable-ppt-v1 `
  --path skills/codeximage-to-editable-ppt-v1
```

安装 V1.2：

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo wiltonesten-web/codeximage-to-editable-ppt-v1-2 `
  --path skills/codeximage-to-editable-ppt-v1-2
```

安装完成后请重启 Codex，使其发现新 Skill。

如需手动安装，只需将 `skills/codeximage-to-editable-ppt-v1-2` 复制到
`%USERPROFILE%\.codex\skills\`。不要将整个仓库根目录作为 Skill 文件夹复制。

## Python 依赖

在仓库根目录执行以下命令安装 Python 软件包：

```powershell
python -m pip install -r skills/codeximage-to-editable-ppt-v1-2/requirements.txt
```

严格的最终渲染与版式质量检查需要在 Windows 上使用 Microsoft PowerPoint。
部分输入处理或回退路径还可能需要 Tesseract OCR、LibreOffice 和 Poppler。

## 典型工作流

可以在 Codex 中使用类似下面的请求调用此 Skill：

```text
使用 $codeximage-to-editable-ppt-v1-2 预检这些幻灯片图片，并将它们重建为可编辑演示文稿。
在启动页面 Worker 之前，先向我展示页面策略和安全的 Worker 数量选择。
```

Skill 会先生成只读预检报告，并请用户确认页面策略和准确的 Worker 数量。
如果执行需要额外的任务树，还必须获得用户明确授权。
它不会在实际采用顺序分批处理时仍宣称实现了完整并发。

## 命令行预检

以下示例用于预检 6 张栅格幻灯片页面：

```powershell
python skills/codeximage-to-editable-ppt-v1-2/scripts/parallel_v1_pages.py preflight `
  slide01.png slide02.png slide03.png slide04.png slide05.png slide06.png `
  --report output.preflight.json `
  --intended-outdir output `
  --available-worker-slots 3 `
  --root-task-count 2 `
  --powerpoint-worker-cap 6
```

预检不会直接完成全部可编辑化重建。请继续通过 Codex 操作，使页面 Worker 能够执行
V1 重建流程及 PowerPoint 质量检查规范。

## 默认成功输出

```text
output/
|-- merged_v1_refined_editable.pptx
|-- parallel_timing_report.json
`-- parallel_v1_finalization.json
```

如果验证失败，系统会保留诊断任务和分片产物供检查，而不会把该次运行视为已完成交付。

## 仓库结构

```text
skills/
`-- codeximage-to-editable-ppt-v1-2/
    |-- SKILL.md
    |-- VERSION
    |-- agents/
    |-- references/
    |-- scripts/
    |-- config.example.yaml
    `-- requirements.txt
```

## 验证

验证 Skill 元数据：

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-creator\scripts\quick_validate.py" `
  skills/codeximage-to-editable-ppt-v1-2
```

在 V1 与 V1.2 Skill 文件夹并列安装的环境中运行随附测试：

```powershell
python -m unittest discover `
  -s skills/codeximage-to-editable-ppt-v1-2/scripts `
  -p "test_*.py" -v
```

## 限制

- 本 Skill 用于编排多页输入；单页任务请直接使用基础 V1 Skill。
- 6、9 等更高并行数量取决于是否有足够的 Codex 任务树和硬件容量，并且必须由用户明确授权创建辅助任务。
- 如果所选 Worker 数量高于硬件安全建议值，必须明确覆盖容量限制。
- 栅格图片无法恢复输入中不存在的隐藏文字、原始矢量图、图表数据、动画或字体文件。
- 严格交付依赖 Microsoft PowerPoint 的高保真渲染和终端验证；仅由脚本生成的基础结果不等同于已完成的可编辑演示文稿。

## 项目状态

本仓库以独立社区项目的形式发布 V1.2.7，与 OpenAI 或 Microsoft 不存在隶属或官方认可关系。

## 许可证

采用 MIT License，详情参见 [LICENSE](LICENSE)。
