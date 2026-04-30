# 焊接工艺技术文档

## TIG焊接技术

**TIG焊**（Tungsten Inert Gas Welding）是一种高质量的焊接方法，广泛用于**不锈钢**和**铝合金**的焊接。

### 主要特点

1. 焊接质量高，焊缝美观
2. 适用于薄板焊接
3. 无飞溅，无烟尘

### 焊接参数

| 参数 | 范围 | 说明 |
|------|------|------|
| 焊接电流 | 50-300A | 根据板厚调整 |
| 焊接电压 | 10-20V | 保持稳定 |
| 焊接速度 | 150-400mm/min | 控制热输入 |

## 常见焊接缺陷

在焊接过程中，可能会出现以下缺陷：

- **气孔**：由于保护气体不足或母材表面污染
- **裂纹**：热裂纹和冷裂纹
- **夹渣**：焊渣未清理干净

## 焊接材料

常用的焊接材料包括：

- 焊丝：ER308L、ER316L
- 保护气体：氩气（Ar）、氦气（He）
- 钨极：WT40、WC20

## 安全注意事项

> 焊接作业时必须佩戴防护面罩和手套

```python
# 示例代码：计算焊接热输入
def calculate_heat_input(current, voltage, speed):
    """
    计算焊接热输入
    :param current: 焊接电流 (A)
    :param voltage: 焊接电压 (V)
    :param speed: 焊接速度 (mm/min)
    :return: 热输入 (kJ/mm)
    """
    return (current * voltage * 60) / (speed * 1000)
```

## 参考资料

更多信息请参考：[焊接技术手册](https://example.com/welding-guide)
