# 宏与函数说明

本文档用于说明当前 VBA 宏文件中的主要入口、主流程和各模块功能。

面向对象：

- 需要维护本项目的人
- 想单独运行某些宏的人
- 需要理解调用关系的人

普通用户建议优先阅读 [`README.md`](./README.md)。

## 1. 主要公开入口

### `FormatThesisToSDUTCM`

主入口。

适合大多数用户直接运行。

默认会执行：

1. 页面设置
2. 标题格式化
3. 正文格式化
4. 摘要和关键词处理
5. 目录处理
6. 参考文献处理
7. 表格处理
8. 图片与图题处理
9. 默认混合页码

### `ApplyMixedPageNumbersByTOC`

单独应用混合页码：

- 目录前使用罗马数字
- 目录后使用阿拉伯数字

### `ApplyArabicPageNumbersOnly`

单独应用全文阿拉伯数字页码。

## 2. 主流程

### `RunSDUTCMFormatting`

内部统一主流程。

这是 `FormatThesisToSDUTCM` 真正调用的执行过程。

主要职责：

- 配置标题、正文和目录标题等 Word 样式定义
- 对全文段落进行一次遍历
- 按样式和大纲级别分别处理题目、一级标题、二级标题、三级标题、正文和代码块段落
- 串联目录、参考文献、图片、页码等模块

## 3. 标题与正文相关

### 旧版整篇扫描宏

- `FormatTitleByHeadingStyle`
- `FormatLevel1Heading`
- `FormatLevel2Heading`
- `FormatLevel3Heading`
- `FormatBodyText`

这些过程会按整篇遍历方式处理格式。

当前总流程主要不再直接依赖它们，但它们仍保留在文件中，适合作为独立工具或历史兼容逻辑参考。

### 当前总流程使用的单段处理宏

- `FormatTitleParagraph`
- `FormatLevel1Paragraph`
- `FormatLevel2Paragraph`
- `FormatLevel3Paragraph`
- `FormatBodyParagraph`
- `FormatCompactParagraph`

当前流程会先配置 Word 样式定义，再用这些单段处理宏覆盖已有段落格式。

它们由 `RunSDUTCMFormatting` 调用。

`FormatCompactParagraph` 用于处理 Pandoc 常见的 `Source Code`、`Code` 代码块样式，只清理缩进并保留原有代码高亮。

`Compact` 不再作为代码块处理，因为 Pandoc 也可能把紧凑列表标记为 `Compact`。

## 4. 摘要与关键词

### `MergeAndFormatAbstract`

负责：

- 识别摘要/关键词中英文标题
- 标题与正文分段时尝试合并
- 自动补冒号
- 调整字体、粗细、缩进

这是当前摘要处理的主过程。

## 5. 目录相关

### `ProcessTableOfContents`

目录总入口。

分两种情况：

1. 已有目录域
   调用 `NormalizeExistingTableOfContents`
2. 仅有“目录”占位
   走插入目录逻辑

### `InsertTableOfContents`

用于在“目录”位置插入新的 Word 目录域。

### `UpdateTableOfContents`

更新已有目录域。

### `NormalizeExistingTableOfContents`

当文档中已经存在目录域时：

- 将目录标题统一为“目录”
- 统一目录标题样式
- 更新目录
- 格式化目录条目

### `FormatTableOfContentsEntries`

统一目录条目的字体和加粗状态。

### `GetFirstTOCField`

返回文档中的第一个目录域。

### `ApplyTOCTitleStyle`

统一应用目录标题样式。

### `ConfigureTOCTitleStyle`

如果文档中存在 `TOC 标题` 样式，则将该样式配置为当前项目的目标样式。

## 6. 参考文献相关

### `ProcessReferencesWithSort`

参考文献总入口。

默认流程：

1. `FormatReferences`
2. `FormatReferenceEntries`

虽然保留了历史名称 `ProcessReferencesWithSort`，当前默认流程已经不再调用 `SortReferences`。

### `FormatReferences`

处理参考文献标题。

### `FormatReferenceEntries`

处理参考文献条目格式，包括悬挂缩进。

如果参考文献后面还有新的一级、二级或三级标题，会给后续标题设置段前分页。

### `SortReferences`

对参考文献条目按字母顺序排序。

当前总流程默认不调用它。这个过程会直接改写参考文献段落，使用前应先备份文档。

### `AutoNumberReferences`

自动为参考文献条目添加编号。

当前总流程默认不调用它。

### `ProcessReferences`

不含排序的参考文献处理入口。

当前总流程默认使用的是 `ProcessReferencesWithSort`。

## 7. 表格相关

### `ApplySingleSpacingToTables`

默认主流程使用的内部过程。按表格整体 `Range` 将单元格段落设为单倍行距，不改动字体、对齐、缩进、边框、列宽或 AutoFit 设置。

### `ProcessTables`

可选的三线表格式化入口，遍历文档中的所有 Word 表格；默认主流程不调用它。

### `ApplyThreeLineTableStyle`

将单个表格处理为居中三线表：

- 表格整体居中
- 单元格内容水平和垂直居中
- 只保留顶线、表头下线、底线
- 清除竖线和内部横线

当前只处理表格本体，不自动处理表题。

## 8. 图片相关

### `ProcessImages`

图片总入口。

内部依次调用：

- `FormatImages`
- `FormatImageCaptions`

### `FormatImages`

处理图片居中。

### `FormatImageCaptions`

识别图题并统一图题格式。

图题识别会先归一化常见空白字符，以兼容 Pandoc 从 Markdown 图片 alt 文本生成的说明段落。

## 9. 页码相关

### `ApplyPageNumbers`

内部页码总入口。

通过模式参数决定：

- 全文阿拉伯数字
- 目录前罗马数字、目录后阿拉伯数字

### `EnsureSectionBreakAfterTableOfContents`

用于确保目录后具备页码模式切换所需的节边界。

当前策略比较保守：

- 如果文档已经有多个节，不再自动新增分节
- 主要目标是避免重复运行时不断增加空白页

### `ApplyArabicPageNumbersToAllSections`

为所有节统一设置阿拉伯数字页码。

### `ApplyMixedPageNumbersBySections`

按节设置混合页码。

### `ClearAllPageNumbers`

清空所有节中的已有页码。

### `EnsureCenteredFooterPageNumber`

确保指定节的页脚中存在居中的页码对象。

### `GetSectionIndexByPosition`

根据文档位置返回所在节编号。

## 10. 辅助函数

### 样式名辅助

- `ZhTitleStyleName`
- `ZhHeadingStyleName`
- `ZhBodyTextStyleName`
- `ZhBodyStyleName`

这些函数用于返回中文样式名，避免直接在逻辑里反复写中文常量。

### 分页辅助

#### `EnsurePageBreakBeforeParagraph`

将某个段落设置为“段前分页”。

适合需要幂等分页的场景，比反复插入物理分页符更稳。

## 11. 当前建议

### 对普通用户

直接运行：

`FormatThesisToSDUTCM`

### 对需要改页码的用户

- 默认混合页码：`FormatThesisToSDUTCM`
- 全文阿拉伯数字：运行后再执行 `ApplyArabicPageNumbersOnly`

### 对维护者

如果后续继续修改，建议优先关注：

1. 目录后分节逻辑
2. 页码幂等性
3. 参考文献标题识别和条目格式化
