---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: ea5569d88ea2058e834b89ecbac77842_6e47d38da50911f180e2525400d4ab64
    ReservedCode1: EsidVJEjlWzoPNqLs4KVOQGLSpjt5t8O2glVivfIKb8e8A/C6Jcyn5gCmlHm+/pkrC0jH5GfTM+iHw/zGYeAz4NXaIraOS9pvze2JPf2iHbvC3NbQ58uFl+sHNb/kVKt1whlp1od+oX3Fc9aJqpd07OS6nCmdKSVIFPvkckL/QWcefP1rRiNmLbe/So=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: ea5569d88ea2058e834b89ecbac77842_6e47d38da50911f180e2525400d4ab64
    ReservedCode2: EsidVJEjlWzoPNqLs4KVOQGLSpjt5t8O2glVivfIKb8e8A/C6Jcyn5gCmlHm+/pkrC0jH5GfTM+iHw/zGYeAz4NXaIraOS9pvze2JPf2iHbvC3NbQ58uFl+sHNb/kVKt1whlp1od+oX3Fc9aJqpd07OS6nCmdKSVIFPvkckL/QWcefP1rRiNmLbe/So=
---

# 慧的菜品选择系统

个人家常菜谱管理小工具：25 道菜谱浏览、餐次筛选、按食材搜索、随机选菜、菜谱查看、新增菜品。

在线访问（GitHub Pages）：`https://<你的用户名>.github.io/<仓库名>/`

## 功能

- 餐次筛选：全部 / 早餐 / 正餐 / 甜品
- 按食材搜索：输入食材关键词（如"豆腐"）即时过滤
- 随机选菜：按餐次随机来一道，选后可记录
- 菜谱详情：步骤、食材清单、抖音链接（如有）
- 新增菜品：电脑本地新增会写入 Excel；手机端新增仅保存在手机浏览器本地

## 两种使用方式

### 手机端（只读为主）

直接打开 GitHub Pages 链接即可使用。菜品数据来自本仓库内置的 HTML 文件，无需安装任何东西。
手机端"新增菜品 / 选菜记录"会保存在手机浏览器本地，不会同步回电脑 Excel。

### 电脑本地（完整功能）

1. 安装 Python 3 并安装依赖：`pip install openpyxl flask`
2. 双击运行 `菜品系统.bat`（自动启动本地服务并打开浏览器）
3. 新增菜品、选菜记录会实时写入 `菜品数据.xlsx`
4. 注意：Excel 打开时保存会失败，先关闭 Excel 再运行

## 文件说明

| 文件 | 说明 |
|------|------|
| index.html | 单文件页面（数据内嵌），电脑/手机通用，双击即可打开 |
| 菜品数据.xlsx | 数据源：25 道菜 + 选菜记录 + 统计报表 |
| 菜品系统服务.py | 本地 Python 服务，负责页面与 Excel 之间的读写中转 |
| sync_records.py | 选菜记录同步脚本 |
| 菜品系统.bat | 一键启动（启动服务 + 自动打开浏览器） |

## 更新发布流程（重要）

本地在 Excel 或页面里改完数据后，要让手机端看到最新数据，需要重新生成 HTML 内嵌数据：

1. 双击 `菜品系统.bat` 即可自动完成「选菜记录同步 → Excel 数据更新到 HTML → 打开最新页面」；也可单独运行 `python update_html.py` 手动同步
2. 提交推送：
   ```
   git add .
   git commit -m "更新菜品数据"
   git push
   ```
3. GitHub Pages 约 1~2 分钟后自动生效

## 部署 GitHub Pages

1. 在 GitHub 创建公开仓库（如 `dish-menu`），把本目录文件推上去
2. 仓库 Settings → Pages → Source 选择 `main` 分支 + `/(root)` → Save
3. 等 1~2 分钟，访问 `https://<用户名>.github.io/<仓库名>/`

> 注意：仓库转 private 会自动停用 Pages；转回 public 后需到 Settings → Pages 重新选择分支保存一次。
*（内容由AI生成，仅供参考）*
