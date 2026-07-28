# C-REITs

中国公募 REITs 地图网站。

在线动态入口：https://wqm12.github.io/C-REITs/

## 当前版本（2026-07-28）

本次发布使用最新的 `REITs_location_updating.xlsx`，页面制作月份由生成器自动适配为 `2026-07`。

- 支持“已上市 REITs / 未上市 REITs”模式切换。
- 右侧统计面板按“基础设施REITs / 商业不动产REITs”分区展示。
- 地图支持省份/城市视图、区域导航和项目搜索。
- `index.html` 内嵌 ECharts 与中国地图资源，不依赖外部 CDN。
- `history/20260701/` 保存上一轮线上版本；`history/20260605/`、`history/20260608/` 及更早版本继续保留在 `history/`。

## 主要文件

- `index.html`：GitHub Pages 动态入口，展示最新地图。
- `reits_map.py`：本地正式生成器。
- `REITs_location_updating.xlsx`：已上市与未上市 REITs 数据源。
- `location.csv`：城市和省份经纬度底表。
- `echarts.min.js`、`china.js`：地图前端依赖。
- `辅助文件/map_source_template.py`：生成器内部模板。
- `地图设计思路与结构说明.md`：地图结构和日常更新说明。
- `history/`：历史发布版本。

## 本地生成

更新 Excel 后运行：

```powershell
python reits_map.py
```

脚本会自动生成 `REITs详细分布图_YYYYMMDD.html`，制作月份自动取当前年月。
