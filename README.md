# KSOA天猫对账自动化

电商财务对账专家Skill，熟练掌握天猫订单从KSOA零售到万里牛出库到支付宝/聚合账户的全链路数据匹配与差异核对。

## 文件说明

| 文件 | 说明 |
|---|---|
| skill-ksoa-tmall-recon.md | Skill完整文档 |
| tmall-recon-data-model.md | 数据材料关联模型 |
| recon_ksoa_tmall_v2.py | 对账主脚本 |
| generate_ksoa_tmall_report_v2.py | Excel报告生成器 |

## 核心功能

1. **匹配链路**：KSOA doc_no → 万里牛系统单号 → 天猫订单号
2. **等式验证**：KSOA小计 = 支付宝到账 + 聚合结算到账
3. **差异分类**：含关闭订单（橙色）/ 正常差异（红色）/ 未匹配（黄色）
4. **合并订单处理**：一个KSOA单据对应多个天猫订单时直接合并核对

## 调用方式

在Claude Code中激活Skill：/KSOA天猫对账
