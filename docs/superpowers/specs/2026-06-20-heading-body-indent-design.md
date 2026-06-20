# 标题与正文缩进修正设计

## 背景

当前样式配置明确将题目和一至三级标题的首行、左侧、右侧缩进设为 0，这是标题无缩进所必需的规则。正文规范要求首行缩进 24 pt，但实现将 `Normal` 和 `First Paragraph` 配置为 0，与 `FORMAT_SPEC.md` 不一致，导致部分首段无缩进。

## 目标

- 题目和一至三级标题继续明确保持首行、左侧、右侧缩进为 0。
- `正文文本`、`正文`、`Normal`、`First Paragraph` 四种正文样式统一首行缩进 24 pt。
- 不判断段落语言，优先满足中文论文正文缩进要求。
- 不新增逐段扫描，不影响激进性能路径。

## 实现

仅修改 `ConfigureSDUTCMStyles` 调用参数：

- `ConfigureBodyStyleIfExists "Normal", 0` 改为 `24`。
- `ConfigureBodyStyleIfExists "First Paragraph", 0` 改为 `24`。

标题样式中的 `.FirstLineIndent = 0`、`.LeftIndent = 0`、`.RightIndent = 0` 保持不变。

## 验证

- 静态测试确认四种正文样式均传入 `24`。
- 静态测试确认题目和标题样式仍明确设置三种缩进为 0。
- 现有性能、参考文献和表格测试继续通过。
