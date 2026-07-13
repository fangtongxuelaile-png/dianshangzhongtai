# 小豚当家影锋风格 BI 看板：在线部署版

## 在线效果

部署成功后，浏览器打开网址即可使用。页面左侧自带「更新数据源」入口，上传最新 Excel 后自动刷新全套看板。

## 推荐部署方式一：Streamlit Community Cloud

1. 把本文件夹上传到 GitHub 仓库。
2. 打开 Streamlit Community Cloud。
3. New app，选择仓库。
4. Main file path 填：`app.py`。
5. 部署完成后复制网址给团队使用。

## 推荐部署方式二：Render

1. 把本文件夹上传到 GitHub 仓库。
2. 打开 Render，New Web Service。
3. 选择该仓库。
4. Build Command：`pip install -r requirements.txt`
5. Start Command：`streamlit run app.py --server.address=0.0.0.0 --server.port=$PORT`
6. 部署完成后即可在线访问。

## 推荐部署方式三：Docker/服务器

```bash
docker build -t xiaotun-yingfeng-bi .
docker run -p 8501:8501 xiaotun-yingfeng-bi
```

然后访问：`http://服务器IP:8501`

## 每日维护流程

1. 打开在线网址。
2. 左侧上传最新 Excel 数据源。
3. 看板自动刷新。
4. 如需留档，可下载离线 HTML 或 CSV。

## 说明

- 上传文件只在当前会话中解析，不会默认保存到服务器。
- 如果要多人共享同一份“最新数据”，可后续增加数据库/对象存储，用于保存最新上传的数据源。
