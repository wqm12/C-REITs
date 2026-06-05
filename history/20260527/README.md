# C-REITs

中国公募 REITs 地图网站。
在线入口：https://wqm12.github.io/C-REITs/

## 当前版本（2026-05-27）

本次发布使用 2026-05-27 更新后的 `REITs_location_updating.xlsx`，新增易方达广西北投高速公路记录，并将最新生成的 `REITs详细分布图_20260527.html` 同步为仓库根目录 `index.html`。

- 页面底部制作日期由 `DISPLAY_DATE = datetime.now().strftime("%Y-%m")` 自动生成，当前显示 `2026-05`。
- 生成批次文件名由 `RUN_DATE_SUFFIX = datetime.now().strftime("%Y%m%d")` 自动生成，当前批次为 `20260527`。
- `history/20260521/` 保存上一轮线上版本，`history/20260506/` 和 `history/20260426/` 继续保留更早历史版本。
- `index.html` 已内嵌 `echarts.min.js` 和 `china.js`，不依赖外部 CDN。

> 注：下方为仓库早期说明，当前版本信息以本节为准。

中国公募 REITs 地图网站。

在线入口：

https://wqm12.github.io/C-REITs/

## 当前版本

当前根目录是 2026-05-21 最新版本，访问 GitHub Pages 时会直接打开最新地图页面。

主要变化：

- 使用 `reits_map.py` 作为最新维护脚本。
- 使用 `REITs详细分布图_20260521.html` 同步生成根目录 `index.html`。
- 地图底部制作日期按运行月份自动生成，本版显示为 `2026-05`。
- 内嵌 `echarts.min.js` 和 `china.js`，减少外部 CDN 依赖，方便国内网络环境直接访问。
- 保留上一轮 GitHub Pages 版本到 `history/20260506/`，并继续保留更早的 `history/20260426/`。

## 文件说明

- `index.html`：GitHub Pages 入口文件，直接展示最新地图。
- `reits_map.py`：生成当前地图的 Python 脚本。
- `REITs_location_updating.xlsx`：当前 REITs 明细数据源。
- `location.csv`：城市和省份经纬度匹配底表。
- `echarts.min.js`：ECharts 前端运行依赖。
- `china.js`：中国地图前端运行依赖。
- `地图设计思路与结构说明.md`：当前地图设计、数据流和发布结构说明。
- `history/20260506/`：上一轮 GitHub Pages 版本的历史快照。
- `history/20260426/`：更早一轮 GitHub 上传版本的历史快照。

## 本地生成方式

在项目目录下运行：

```bash
python reits_map.py
```

脚本会读取 `REITs_location_updating.xlsx` 和 `location.csv`，生成日期版 HTML。发布到 GitHub Pages 时，将最新 HTML 同步为仓库根目录的 `index.html`。

## 数据格式

`项目分类` 字段采用以下格式维护：

```text
{城市1资产A,城市1资产B}|{城市2资产A}
```

每个 `{}` 对应一个城市，不同城市之间用 `|` 分隔，城市顺序与 `项目城市` 字段保持一致。
