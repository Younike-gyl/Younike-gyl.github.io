---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 16010dfaaa626db7193d8d6a2a994cf5_b0fd6b408a1911f1be80525400f8a581
    ReservedCode1: IKeHaXKhchv8takD+X2dh84ftwDbCCAvmJV2wnt2Qh/iXpP6T9WFwlRvmcun8y+tZGk1KMxZ1a+IPKFvkgh1u3x399wDirH5W6TTpVRZp/+60NLobnvcwMBleo+8JKjXKvcpO3x0Z5i705xPzkANvATIAznzH6173wo/T0YVjITi8gfm+0cWhuEbfOM=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 16010dfaaa626db7193d8d6a2a994cf5_b0fd6b408a1911f1be80525400f8a581
    ReservedCode2: IKeHaXKhchv8takD+X2dh84ftwDbCCAvmJV2wnt2Qh/iXpP6T9WFwlRvmcun8y+tZGk1KMxZ1a+IPKFvkgh1u3x399wDirH5W6TTpVRZp/+60NLobnvcwMBleo+8JKjXKvcpO3x0Z5i705xPzkANvATIAznzH6173wo/T0YVjITi8gfm+0cWhuEbfOM=
---

<!--
Markdown 语法参考（本文示例涵盖以下要素）：

# 一级标题 → 文章标题
## 二级标题 → 章节
### 三级标题 → 小节

**粗体** → 函数名、重点
*斜体* → 参数说明

内联代码：`函数名()`
代码块：多行公式示例

表格：函数速查表
无序列表、有序列表

> 提示/技巧引用块

--- 分隔线
-->

# Excel 常用函数速查

## 1. 逻辑判断

### IF 函数

```
=IF(条件, 真值, 假值)
```

**示例**：判断质检是否合格

```
=IF(D2>=60, "合格", "不合格")
```

### IFS 多条件判断（Excel 2019+）

```
=IFS(条件1, 结果1, 条件2, 结果2, ...)
```

**示例**：按分数划分等级

```
=IFS(B2>=90, "A", B2>=75, "B", B2>=60, "C", TRUE, "D")
```

## 2. 查找匹配

| 函数 | 用途 | 语法要点 |
|------|------|----------|
| `VLOOKUP` | 垂直查找 | `=VLOOKUP(查找值, 表格范围, 返回列号, 0)` |
| `XLOOKUP` | 现代查找（365） | `=XLOOKUP(查找值, 查找列, 返回列)` |
| `INDEX+MATCH` | 灵活匹配 | `=INDEX(返回列, MATCH(查找值, 查找列, 0))` |

### VLOOKUP 实战：根据物料编码查物料名称

```
=VLOOKUP(A2, 物料主数据!$A$2:$D$500, 2, 0)
```

> *技巧*：第四个参数 **0** 代表精确匹配，**1** 代表近似匹配（要求查找列升序排列）。

### XLOOKUP 实战（推荐）

```
=XLOOKUP(A2, 物料主数据!A:A, 物料主数据!B:B, "未找到")
```

优势：不需要数第几列，支持反向查找，自带容错。

## 3. 条件统计

| 函数 | 用途 | 示例 |
|------|------|------|
| `SUMIF` | 条件求和 | `=SUMIF(A:A, "合格", C:C)` |
| `SUMIFS` | 多条件求和 | `=SUMIFS(C:C, A:A, "合格", B:B, ">=100")` |
| `COUNTIF` | 条件计数 | `=COUNTIF(D:D, "不合格")` |
| `COUNTIFS` | 多条件计数 | `=COUNTIFS(A:A, "供应商A", D:D, "不合格")` |
| `AVERAGEIF` | 条件平均 | `=AVERAGEIF(B:B, ">0", C:C)` |

## 4. 日期与时间

```
=TODAY()                → 当前日期
=NOW()                  → 当前日期+时间
=YEAR(A2)               → 提取年份
=MONTH(A2)              → 提取月份
=DATEDIF(A2, B2, "D")   → 计算间隔天数
=EOMONTH(A2, 0)         → 当月最后一天
=NETWORKDAYS(A2, B2)    → 工作日天数
```

**实战**：计算订单履约天数（自然日）

```
=DATEDIF(下单日期, 签收日期, "D")
```

## 5. 文本处理

```
=LEFT(A2, 4)              → 取左边 4 个字符
=RIGHT(A2, 3)             → 取右边 3 个字符
=MID(A2, 3, 5)            → 从第 3 位取 5 个字符
=CONCAT(A2, "-", B2)      → 拼接文本（365 为 CONCAT，旧版用 CONCATENATE）
=TEXT(A2, "00000")        → 数字格式化（如 PO 单号补零）
=TRIM(A2)                 → 去除首尾空格
```

**实战**：从完整库位编码 `A-01-03` 中提取货架编号

```
=LEFT(A2, FIND("-", A2)-1)
```

## 6. 快速组合技巧

### 数据透视表快捷操作

| 操作 | 快捷键 |
|------|:------:|
| 插入数据透视表 | `Alt + N + V` |
| 刷新全部数据 | `Ctrl + Alt + F5` |
| 插入表格（Ctrl+T 套用格式） | `Ctrl + T` |
| 自动求和 | `Alt + =` |

### 条件格式速设

> 供应链场景常用：对超时订单自动标红。

1. 选中数据区域
2. 开始 → 条件格式 → 新建规则
3. 公式：`=$E2>TODAY()`（假设 E 列为期望日期）
4. 设置填充色为浅红色

## 7. 常见坑点

1. **VLOOKUP 无法向左查找** → 改用 `XLOOKUP` 或 `INDEX+MATCH`
2. **合并单元格导致公式错误** → 尽量避免使用合并单元格存放数据
3. **文本格式数字无法求和** → 选中 → 数据 → 分列 → 完成
4. **跨表引用路径变化** → 使用 `Ctrl+H` 替换更新引用路径

---

*最后更新：2026-07-28 | 版本：V1.0*
*（内容由AI生成，仅供参考）*
