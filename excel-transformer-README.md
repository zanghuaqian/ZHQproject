# Excel Transformer Skill

一个强大的Excel文档转换生成器技能，支持根据用户上传的Excel文档或本地文件夹中的xlsx文件，以及指定的要求、格式规范，生成符合要求的新Excel文档。

## 功能特性

- ✅ 支持读取本地文件夹中的指定xlsx文件
- ✅ 支持数据转换、格式设置、公式计算
- ✅ 支持多表合并、数据透视
- ✅ 特别支持生成"盛意旺采购商城对账单"标准格式
- ✅ 自动数据完整性检查
- ✅ 智能字段映射
- ✅ 自动公式计算和格式设置

## 安装

将 `excel-transformer.skill` 文件复制到你的技能目录即可使用。

## 使用方法

### 基本用法

用户只需提供：
1. Excel文件路径（本地文件路径）
2. 输出格式要求

技能会自动：
- 读取和分析文件
- 映射字段
- 生成符合要求的Excel文档

### 对账单生成

使用提供的脚本快速生成对账单：

```bash
python scripts/generate_statement.py <原数据文件> <交易对账单文件> <月份> [输出文件]
```

示例：
```bash
python scripts/generate_statement.py data.xlsx reference.xlsx 1
```

## 文件结构

```
excel-transformer/
├── SKILL.md                    # 主技能文档
├── scripts/
│   └── generate_statement.py   # 对账单生成脚本
└── references/
    ├── common-patterns.md      # 常见转换模式
    ├── file-reading.md         # 文件读取指南
    ├── statement-format.md     # 对账单格式规范
    └── statement-mapping.md    # 对账单字段映射规则
```

## 依赖

- pandas
- openpyxl

安装依赖：
```bash
pip install pandas openpyxl
```

## 许可证

请查看技能包中的LICENSE文件（如果有）。

## 更新日志

### v1.0.0
- 初始版本
- 支持基础Excel转换功能
- 支持盛意旺采购商城对账单生成
- 完整的字段映射规则
- 自动退款处理
- 支付渠道自动获取
