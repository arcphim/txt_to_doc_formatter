# 配置文件说明

## title_font (标题字体设置)
- `name`: "方正小标宋_GBK"  // 标题字体名称
- `size`: 18  // 标题字体大小
- `alignment`: "center"  // 标题对齐方式 (center: 居中)
- `bold`: false  // 是否加粗 (false: 不加粗)

## body_font (正文字体设置)
- `name`: "方正仿宋_GBK"  // 正文字体名称
- `size`: 16  // 正文字体大小
- `digit_font`: "Times New Roman"  // 数字字体
- `alignment`: "justify"  // 正文对齐方式 (justify: 两端对齐)

## output_dirs (输出目录设置)
- `default_dirs`: ["D:/my工作台/missions"]  // 默认输出目录列表
- `current_dir`: "C:/Users/yuan/OneDrive/桌面"  // 当前输出目录

## spacing (行距设置)
- `line_spacing`: 28  // 行距值
- `line_spacing_rule`: "FIXED"  // 行距规则 (FIXED: 固定值)

## margins (页边距设置)
- `top`: 3.0  // 上边距
- `bottom`: 3.0  // 下边距
- `left`: 2.6  // 左边距
- `right`: 2.6  // 右边距
- `unit`: "cm"  // 边距单位

## heading_levels (标题级别配置)

### 一级标题
- `level`: 1
- `format`: "一、"  // 标题格式
- `font`: "方正黑体_GBK"  // 字体
- `size`: 16  // 字体大小
- `bold`: false  // 是否加粗
- `indent`: 2  // 缩进字符数
- `punctuation`: false  // 是否包含标点
- `new_page`: false  // 是否新起一页
- `standalone_line`: true  // 是否独立成行

### 二级标题
- `level`: 2
- `format`: "（一）"  // 标题格式
- `font`: "方正楷体_GBK"  // 字体
- `size`: 16  // 字体大小
- `bold`: true  // 是否加粗
- `indent`: 2  // 缩进字符数
- `punctuation`: false  // 是否包含标点
- `standalone_line`: true  // 是否独立成行

### 三级标题
- `level`: 3
- `format`: "1. "  // 标题格式
- `font`: "方正仿宋_GBK"  // 字体
- `size`: 16  // 字体大小
- `bold`: false  // 是否加粗
- `indent`: 2  // 缩进字符数
- `punctuation`: true  // 是否包含标点
- `standalone_line`: true  // 是否独立成行

### 四级标题
- `level`: 4
- `format`: "（1）"  // 标题格式
- `font`: "方正仿宋_GBK"  // 字体
- `size`: 16  // 字体大小
- `bold`: false  // 是否加粗
- `indent`: 2  // 缩进字符数
- `punctuation`: true  // 是否包含标点
- `standalone_line`: false  // 是否独立成行

## page_number (页码设置)
- `position`: "bottom"  // 页码位置 (bottom: 底部)
- `alignment`: "center"  // 页码对齐方式 (center: 居中)