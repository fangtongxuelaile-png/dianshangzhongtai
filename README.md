# 小豚当家销售经营BI看板

基于 Streamlit 的电商销售经营 BI 看板，支持 Excel 数据源上传、多维度分析、趋势可视化。

## 功能

- **经营总览**：支付金额、支付件数、支付买家、访客、支付转化率、客单价、退款率
- **推广分析**：推广投放数据多维度分析
- **时间段对比**：同比/环比对比
- **趋势分析**：日/周/月趋势图
- **智能诊断**：异常检测、TOP商品分析
- **透视表分析**：多维度交叉分析
- **目标达成**：月度目标达成率追踪

## 部署

### Streamlit Cloud
1. Fork 本仓库
2. 登录 [Streamlit Cloud](https://share.streamlit.io/)
3. 创建新 App，选择本仓库 + `app.py`
4. 访问生成的 URL

### Render
```bash
pip install -r requirements.txt
streamlit run app.py
```

### Docker
```bash
docker build -t xiaotun-bi .
docker run -p 8501:8501 xiaotun-bi
```

## 登录
| 账号 | 密码 | 角色 | 过期 |
|------|------|------|------|
| `gwell` | `xiaotun666` | 管理员（admin） | 永久 |
| `xiaotun` | `xiaotun2026` | 浏览者（viewer） | 2027-12-31 |

> 部署在 Streamlit Cloud：https://xiaotunbi.streamlit.app

## 数据源
支持上传 Excel 数据源，包含：
- 天猫数据源
- 京东抖音数据源

核心字段：统计日期、渠道、店铺、品类、型号、款式、商品名称、商品访客数、支付买家数、支付件数、支付金额、成功退款金额等。
