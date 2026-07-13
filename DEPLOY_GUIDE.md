# XiaotunBI 部署指南 — 手机热点 + GitHub + Streamlit Cloud

> **前置条件**：手机开热点 → 电脑连上热点 → 确认能打开 github.com

---

## 第一步：创建 GitHub 仓库（约2分钟）

1. 浏览器打开 **https://github.com**
2. 登录账号 **MarkLv2026**
3. 点击左上角 **"+" → "New repository"**
4. 填写：
   - Repository name: `xiaotunbi`
   - Description: `小豚当家BI看板`
   - **选 Private（私有）**
5. **不要勾选** Add a README / .gitignore / license
6. 点 **Create repository**

## 第二步：创建访问令牌 Token（约2分钟）

1. 打开 **https://github.com/settings/tokens** （或从右上角头像→Settings→Developer settings→Personal access tokens→Tokens (classic)）
2. 点击 **"Generate new token (classic)"**
3. 设置：
   - Note: `streamlit-deploy`
   - Expiration: 选 **90 days**
   - 权限勾选：**repo**（整个打钩）
4. 点 **Generate token**
5. **立即复制token**（只显示一次！格式类似 `ghp_xxxxxxxxxxxx`）

## 第三步：推送代码到 GitHub（我来帮你执行）

拿到 token 后，**发给我**，我直接在电脑上执行推送命令。

或者你自己在命令行执行（复制粘贴）：

```bash
cd c:\Users\Gwell\WorkBuddy\20260515174015
git remote add https-origin https://MarkLv2026:<你的TOKEN>@github.com/MarkLv2026/xiaotunbi.git
git push https-origin main
```

> ⚠️ 把 `<你的TOKEN>` 替换成第二步复制的完整 token

## 第四步：部署到 Streamlit Cloud（约3分钟）

1. 打开 **https://share.streamlit.io**
2. 用 GitHub 账号登录（点 "Sign up" 或 "Log in" → 选 GitHub）
3. 点 **"Deploy an app"** 或 **"New app"**
4. 填写：
   - Repository: `MarkLv2026/xiaotunbi`
   - Branch: `main`
   - Main file path: `./app.py`
5. 点 **"Deploy!"**
6. 等待 1-2 分钟，看到 **"Your app is live!"** 就成功了！🎉

## 第五步：使用

- 复制生成的网址，发给任何人都能打开
- 每个人上传自己的 Excel 数据，数据互不干扰
- 不需要你电脑开着，24小时在线

---

## 常见问题

| 问题 | 解决方法 |
|------|---------|
| github.com 打不开 | 确认已连接手机热点，换个浏览器试试 |
| 登录后页面空白 | 等10秒刷新，或换 Chrome 浏览器 |
| Token 忘记复制 | 回到 tokens 页面删除旧的，重新生成一个 |
| 推送报错 403 | Token 错误或过期，重新生成 |
| Streamlit 登录跳转失败 | 确保 GitHub 已登录，直接访问 share.streamlit.io |

## 全部完成后的效果

✅ 公网可访问的网址（如 `xxxxx.app.streamlit.app`）
✅ 多人同时在线，各自上传各自数据
✅ 免费永久运行，不需要服务器、不需要开电脑
