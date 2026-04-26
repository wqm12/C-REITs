# C-REITs

本项目展示中国大陆公募 REITs 的最新地图分布，并提供全国 / 城市双层缩放视图、区域导航、搜索、统计面板等交互功能。

## 在线访问

- GitHub 仓库: https://github.com/wqm12/C-REITs
- 网站入口: https://wqm12.github.io/C-REITs/

## 当前版本说明

当前版本已更新为最新地图页面，主要特性包括：

- 基于 `REITs_location_updating.xlsx` 和 `location.csv` 自动生成地图
- 全国省份层与城市层双模式渲染
- 缩放阈值切换，缩回后可恢复到稳定的全国视图
- 顶部区域导航、左侧搜索、右侧资产统计面板
- 适配新的项目分类格式：`{城市项目组}|{城市项目组}`

## 主要文件

- `index.html`：GitHub Pages 网站入口，始终同步最新地图 HTML
- `Full Code.ipynb`：历史完整版本参考
- `reits_map_updated.py`：当前可运行主脚本
- `REITs地图更新版.ipynb`：当前同步 notebook 版本

## 更新方式

每次本地地图更新后，只需同步仓库根目录的 `index.html`，GitHub Pages 即可对外展示最新版本。
