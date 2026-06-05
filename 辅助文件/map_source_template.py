import pandas as pd
from pyecharts.charts import Geo
from pyecharts import options as opts
from pyecharts.globals import ChartType
from pyecharts.commons.utils import JsCode
import json
import os
import re
import math
from datetime import datetime
from pathlib import Path

# ========= 区域配置 =========
CLUSTER_INFO = {
    "京津冀": {"center": [116.4, 39.5], "zoom": 6, "provinces": ["北京", "天津", "河北"]},
    "长三角": {"center": [120.5, 31.5], "zoom": 6, "provinces": ["上海", "江苏", "浙江", "安徽"]},
    "粤港澳大湾区": {"center": [113.5, 22.8], "zoom": 7, "provinces": ["广东", "香港", "澳门"]},
    "成渝": {"center": [105.5, 30.0], "zoom": 6, "provinces": ["四川", "重庆"]},
    "长江中游": {"center": [114.3, 30.6], "zoom": 6, "provinces": ["湖北", "湖南", "江西"]},
}

# ========= 统一配色方案 =========
VISUAL_MAP_COLORS = [
    "#4575b4", "#3C9BC9", "#74add1", "#95CBC9", "#ADC9AB",
    "#E2C89A", "#FBC08A", "#FAA26F", "#F97F5F", "#F46d43", "#d73027",
]


def is_missing(value):
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    text = str(value).strip()
    return text == "" or text.lower() == "nan"


def find_project_root(anchor_path, required_files=None):
    required = tuple(required_files or ("REITs_location_updating.xlsx", "location.csv"))
    anchor = Path(anchor_path).resolve()
    current = anchor if anchor.is_dir() else anchor.parent

    for candidate in (current, *current.parents):
        if all((candidate / name).exists() for name in required):
            return candidate

    return current


def split_place_names(text):
    if is_missing(text):
        return []
    return [part.strip() for part in re.split(r"[，,]", str(text)) if part.strip()]


def split_asset_items(group_text):
    if is_missing(group_text):
        return []

    text = str(group_text).strip()
    items = []
    current = []
    depth = 0

    for ch in text:
        if ch == "(":
            depth += 1
            current.append(ch)
            continue
        if ch == ")":
            depth = max(depth - 1, 0)
            current.append(ch)
            continue
        if ch in {"，", ","} and depth == 0:
            item = "".join(current).strip()
            if item:
                items.append(item)
            current = []
            continue
        current.append(ch)

    tail = "".join(current).strip()
    if tail:
        items.append(tail)

    return items


def parse_brace_groups(text):
    if is_missing(text):
        return []

    source = str(text).strip()
    groups = []
    current = []
    depth = 0

    for ch in source:
        if ch == "{":
            if depth > 0:
                current.append(ch)
            depth += 1
            continue
        if ch == "}":
            depth -= 1
            if depth > 0:
                current.append(ch)
            elif depth == 0:
                groups.append("".join(current).strip())
                current = []
            continue
        if depth > 0:
            current.append(ch)

    return groups


def parse_underlying_assets(project_classification, cities):
    """
    解析项目分类中的底层资产。
    现格式示例: "{资产1，资产2}|{资产3}"，每个 {} 对应一个城市。
    """
    clean_cities = [city.strip() for city in cities if str(city).strip()]
    if not clean_cities:
        return {}

    groups = parse_brace_groups(project_classification)
    if not groups:
        legacy_text = str(project_classification).strip()
        if legacy_text.startswith("(") and legacy_text.endswith(")"):
            legacy_text = legacy_text[1:-1]
        groups = [legacy_text]

    result = {}
    for index, city in enumerate(clean_cities):
        if len(clean_cities) == 1 and groups:
            result[city] = split_asset_items(groups[0])
            continue
        group_text = groups[index] if index < len(groups) else ""
        result[city] = split_asset_items(group_text)
    return result


def build_location_lookup(df_loc):
    working = df_loc.dropna(subset=["name", "longitude", "latitude"]).copy()
    working["name"] = working["name"].astype(str).str.strip()
    working = working.drop_duplicates(subset="name", keep="first")
    return working.set_index("name")[["longitude", "latitude"]].to_dict("index")


def build_coord_string(place_text, loc_dict):
    names = split_place_names(place_text)
    coords = []
    for name in names:
        if name not in loc_dict:
            continue
        lon = loc_dict[name]["longitude"]
        lat = loc_dict[name]["latitude"]
        coords.append(f"({lon},{lat})")
    return ",".join(coords) if coords else None


def parse_coord_pairs(coord_text):
    if is_missing(coord_text):
        return []

    raw = str(coord_text).replace("),(", "|").strip("()")
    if not raw:
        return []

    pairs = []
    for item in raw.split("|"):
        parts = [part.strip() for part in item.split(",")]
        if len(parts) < 2:
            continue
        try:
            pairs.append((float(parts[0]), float(parts[1])))
        except ValueError:
            continue
    return pairs


def count_coord_groups(coord_text):
    if is_missing(coord_text):
        return 0
    return len(parse_coord_pairs(coord_text))


# ========= 1. 数据预处理 =========
ANCHOR_PATH = Path(__file__).resolve() if "__file__" in globals() else Path.cwd()
PROJECT_DIR = find_project_root(ANCHOR_PATH)
DATA_FILE = PROJECT_DIR / "REITs_location_updating.xlsx"
LOCATION_FILE = PROJECT_DIR / "location.csv"
RUN_DATE_SUFFIX = datetime.now().strftime("%Y%m%d")
DISPLAY_DATE = datetime.now().strftime("%Y-%m")
PROCESS_ROOT = PROJECT_DIR / "process_files"
PROCESS_ROOT.mkdir(exist_ok=True)
PROCESS_DIR = PROCESS_ROOT / f"process_{RUN_DATE_SUFFIX}_local_assets"
PROCESS_DIR.mkdir(exist_ok=True)
TEMP_HTML_PATH = PROCESS_DIR / f"REITs详细分布图_{RUN_DATE_SUFFIX}temp.html"
FINAL_HTML_PATH = PROJECT_DIR / f"REITs详细分布图_{RUN_DATE_SUFFIX}.html"
LOCAL_ASSET_FILES = {
    "https://assets.pyecharts.org/assets/v6/echarts.min.js": PROJECT_DIR / "echarts.min.js",
    "https://assets.pyecharts.org/assets/v6/maps/china.js": PROJECT_DIR / "china.js",
}


def embed_local_js_assets(html_content):
    for remote_asset_url, local_asset_path in LOCAL_ASSET_FILES.items():
        if not local_asset_path.exists():
            raise FileNotFoundError(f"缺少本地 JS 资源: {local_asset_path}")

        js_content = local_asset_path.read_text(encoding="utf-8")
        js_content = js_content.replace("</script>", "<\\/script>")
        inline_script = f'<script type="text/javascript">\n{js_content}\n</script>'
        script_pattern = (
            r'<script\s+type="text/javascript"\s+src="'
            + re.escape(remote_asset_url)
            + r'"></script>'
        )
        html_content, replacement_count = re.subn(
            script_pattern,
            lambda _: inline_script,
            html_content,
            count=1,
        )
        if replacement_count != 1:
            raise ValueError(f"未找到待内嵌的远程 JS 引用: {remote_asset_url}")

    return html_content

df = pd.read_excel(DATA_FILE)
df_loc = pd.read_csv(LOCATION_FILE)
loc_dict = build_location_lookup(df_loc)

df["城市经纬度"] = df["项目城市"].apply(lambda x: build_coord_string(x, loc_dict))
df["省份经纬度"] = df["项目省份"].apply(lambda x: build_coord_string(x, loc_dict))

city_coord_mismatches = []
province_coord_mismatches = []

for _, row in df.iterrows():
    city_input_count = len(split_place_names(row["项目城市"]))
    city_output_count = count_coord_groups(row["城市经纬度"])
    if city_input_count != city_output_count:
        city_coord_mismatches.append({
            "证券代码": row["证券代码"],
            "项目城市": row["项目城市"],
            "城市经纬度": row["城市经纬度"],
        })

    province_input_count = len(split_place_names(row["项目省份"]))
    province_output_count = count_coord_groups(row["省份经纬度"])
    if province_input_count != province_output_count:
        province_coord_mismatches.append({
            "证券代码": row["证券代码"],
            "项目省份": row["项目省份"],
            "省份经纬度": row["省份经纬度"],
        })

if city_coord_mismatches:
    print("以下记录的城市经纬度匹配不完整：")
    print(pd.DataFrame(city_coord_mismatches))

if province_coord_mismatches:
    print("以下记录的省份经纬度匹配不完整：")
    print(pd.DataFrame(province_coord_mismatches))

df["资产类型"] = df["资产类型"].str.replace("基础设施", "", regex=False)
rename_map = {
    "园区": "产业园",
    "交通": "交通设施",
    "消费": "消费商业",
    "保障性租赁住房": "保租房",
    "仓储物流": "物流仓储",
    "新型": "数据中心",
    "市政": "市政公用",
    "水利": "水利",
}
df["资产类型"] = df["资产类型"].replace(rename_map)

total_reits_count = len(df["证券名称"].dropna().unique())

# ========= 2. 解析省份和城市数据 =========
province_rows = []
city_rows = []
asset_details = []

for _, row in df.iterrows():
    provinces = split_place_names(row["项目省份"])
    province_coords = parse_coord_pairs(row["省份经纬度"])

    cities = split_place_names(row["项目城市"])
    city_coords = parse_coord_pairs(row["城市经纬度"])
    city_assets_map = parse_underlying_assets(row["项目分类"], cities)

    for index, province_name in enumerate(provinces):
        if index >= len(province_coords):
            continue
        lon, lat = province_coords[index]
        province_rows.append({
            "province": province_name,
            "lon": lon,
            "lat": lat,
            "security": row["证券名称"],
            "asset_type": row["资产类型"],
        })

    for index, city_name in enumerate(cities):
        if index >= len(city_coords):
            continue
        lon, lat = city_coords[index]
        underlying_assets = city_assets_map.get(city_name, [])
        city_rows.append({
            "city": city_name,
            "project": "、".join(underlying_assets) if underlying_assets else row["项目名称"],
            "lon": lon,
            "lat": lat,
            "security": row["证券名称"],
            "asset_type": row["资产类型"],
            "underlying_assets": underlying_assets,
        })
        asset_details.append({
            "security": row["证券名称"],
            "asset_type": row["资产类型"],
            "city": city_name,
            "underlying_assets": underlying_assets,
        })

province_df = pd.DataFrame(province_rows)
city_df = pd.DataFrame(city_rows)

# ========= 3. 汇总省份统计 =========
province_stats = {}
for _, row in province_df.iterrows():
    province = row["province"]
    security_info = f"{row['security']}({row['asset_type']})"

    if province not in province_stats:
        province_stats[province] = {"lon": row["lon"], "lat": row["lat"], "securities": []}

    if security_info not in province_stats[province]["securities"]:
        province_stats[province]["securities"].append(security_info)

province_data_pair = []
max_province_count = 0
for province, info in province_stats.items():
    count = len(info["securities"])
    max_province_count = max(max_province_count, count)
    security_list = sorted(info["securities"])
    security_list_str = "<br/>".join([f"• {item}" for item in security_list])
    tooltip_text = f"<b>{province}(共有 {count} 个REITs项目)</b><br/>{security_list_str}"
    province_data_pair.append({
        "name": province,
        "value": [info["lon"], info["lat"], count, tooltip_text],
    })

# ========= 4. 汇总城市统计 =========
city_stats = {}
all_securities_list = []

for _, row in city_df.iterrows():
    city = row["city"]
    security_info = f"{row['security']}({row['asset_type']})"

    if city not in city_stats:
        city_stats[city] = {"lon": row["lon"], "lat": row["lat"], "securities": []}

    if security_info not in city_stats[city]["securities"]:
        city_stats[city]["securities"].append(security_info)

    all_securities_list.append({
        "name": row["security"],
        "full_name": security_info,
        "city": city,
        "lon": row["lon"],
        "lat": row["lat"],
        "asset_type": row["asset_type"],
    })

city_data_pair = []
max_city_count = 0
for city, info in city_stats.items():
    count = len(info["securities"])
    max_city_count = max(max_city_count, count)
    security_list_str = "<br/>".join([f"• {item}" for item in info["securities"]])
    tooltip_text = f"<b>{city}(共有 {count} 个REITs项目)</b><br/>{security_list_str}"
    city_data_pair.append({
        "name": city,
        "value": [info["lon"], info["lat"], count, tooltip_text],
    })

# ========= 构建资产分类层次结构(带排序) =========
asset_hierarchy = {}
for detail in asset_details:
    asset_type = detail["asset_type"]
    security = detail["security"]
    city = detail["city"]
    underlying = detail["underlying_assets"]

    if asset_type not in asset_hierarchy:
        asset_hierarchy[asset_type] = {}
    if security not in asset_hierarchy[asset_type]:
        asset_hierarchy[asset_type][security] = {}
    if city not in asset_hierarchy[asset_type][security]:
        asset_hierarchy[asset_type][security][city] = []

    for asset in underlying:
        if asset not in asset_hierarchy[asset_type][security][city]:
            asset_hierarchy[asset_type][security][city].append(asset)

asset_hierarchy_sorted = []
for asset_type in asset_hierarchy:
    reits_list = []
    total_cities = 0
    unique_reits = set()

    for reit in asset_hierarchy[asset_type]:
        unique_reits.add(reit)
        cities_list = []
        city_count = len(asset_hierarchy[asset_type][reit])
        total_cities += city_count

        for city in asset_hierarchy[asset_type][reit]:
            cities_list.append({
                "city": city,
                "assets": asset_hierarchy[asset_type][reit][city],
                "asset_count": len(asset_hierarchy[asset_type][reit][city]),
            })

        cities_list.sort(key=lambda item: item["asset_count"], reverse=True)

        reits_list.append({
            "reit": reit,
            "cities": cities_list,
            "city_count": city_count,
        })

    reits_list.sort(key=lambda item: item["city_count"], reverse=True)

    asset_hierarchy_sorted.append({
        "asset_type": asset_type,
        "reits": reits_list,
        "total_cities": total_cities,
        "reits_count": len(unique_reits),
    })

asset_hierarchy_sorted.sort(key=lambda item: item["reits_count"], reverse=True)
asset_hierarchy_json = json.dumps(asset_hierarchy_sorted, ensure_ascii=False)

# ========= 5. 创建 Geo 对象 =========
geo = Geo(init_opts=opts.InitOpts(width="100%", height="100vh", bg_color="#f4f4f4"))

geo.add_schema(
    maptype="china",
    itemstyle_opts=opts.ItemStyleOpts(color="#d0e1f9", border_color="#ffffff"),
    zoom=1.65,
    center=[108, 35],
    label_opts=opts.LabelOpts(is_show=False)
)

for province, info in province_stats.items():
    geo.add_coordinate(province, info["lon"], info["lat"])

for city, info in city_stats.items():
    geo.add_coordinate(city, info["lon"], info["lat"])

geo.add(
    series_name="中国REITs分布",
    data_pair=[(x["name"], x["value"][2]) for x in province_data_pair],
    type_=ChartType.EFFECT_SCATTER,
    symbol_size=13,
    label_opts=opts.LabelOpts(is_show=True, formatter="{b}", position="right", font_size=12, color="#1f2d3d", opacity=1, font_weight="600", text_border_color="#ffffff", text_border_width=2)
)

province_formatter_js = """
function (params) {
    var province_info = %s;
    for (var i=0; i<province_info.length; i++) {
        if (province_info[i].name === params.name) {
            return province_info[i].value[3];
        }
    }
    return params.name;
}
""" % str(province_data_pair)

geo.set_global_opts(
    legend_opts=opts.LegendOpts(pos_top="10", pos_left="center"),
    visualmap_opts=opts.VisualMapOpts(
        min_=1, max_=max_province_count, pos_left="10", pos_bottom="10",
        is_piecewise=True,
        pieces=[{"min": i, "max": i, "label": str(i)} for i in range(1, min(11, max_province_count + 1))] +
               ([{"min": 11, "max": max_province_count, "label": f"11+"}] if max_province_count > 10 else []),
        range_color=VISUAL_MAP_COLORS,
        textstyle_opts=opts.TextStyleOpts(font_size=12)
    ),
    tooltip_opts=opts.TooltipOpts(trigger="item", formatter=JsCode(province_formatter_js)),
    graphic_opts=[
        opts.GraphicGroup(
            graphic_item=opts.GraphicItem(left="center", bottom="10"),
            children=[
                opts.GraphicText(
                    graphic_item=opts.GraphicItem(left="center", bottom="10"),
                    graphic_textstyle_opts=opts.GraphicTextStyleOpts(
                        text=f"数据来源: Wind, iFind | 原创制作: Qimai | 制作日期: {DISPLAY_DATE}",
                        font="12px Arial",
                        graphic_basicstyle_opts=opts.GraphicBasicStyleOpts(fill="#666666")
                    )
                )
            ]
        )
    ]
)

# ========= 导出与 HTML 增强 =========
geo.render(str(TEMP_HTML_PATH))

with open(TEMP_HTML_PATH, "r", encoding="utf-8") as f:
    html_content = f.read()

province_data_json = json.dumps(province_data_pair, ensure_ascii=False)
city_data_json = json.dumps(city_data_pair, ensure_ascii=False)
securities_data_json = json.dumps(all_securities_list, ensure_ascii=False)
visual_map_colors_json = json.dumps(VISUAL_MAP_COLORS)

all_coords = {}
for province, info in province_stats.items():
    all_coords[province] = [info["lon"], info["lat"]]
for city, info in city_stats.items():
    all_coords[city] = [info["lon"], info["lat"]]
all_coords_json = json.dumps(all_coords, ensure_ascii=False)

navigation_and_search_html = f"""
<style>
html, body {{
    margin: 0;
    padding: 0;
    width: 100%;
    height: 100%;
    overflow: hidden;
    font-family: "Microsoft YaHei", sans-serif;
}}

#container, .chart-container {{
    width: 100vw !important;
    height: 100vh !important;
    position: fixed;
    top: 0;
    left: 0;
}}

.nav-container {{ position: absolute; top: 20px; left: 20px; z-index: 1000; display: flex; gap: 8px; flex-wrap: wrap; max-width: 600px; }}
.nav-btn {{ padding: 6px 12px; border-radius: 4px; border: 2px solid; background: white; cursor: pointer; font-weight: bold; font-size: 13px; }}
.btn-china {{ border-color: #3498db; color: #3498db; }}
.btn-jingjinji {{ border-color: #e74c3c; color: #e74c3c; }}
.btn-changsan {{ border-color: #2ecc71; color: #2ecc71; }}
.btn-wanqu {{ border-color: #f39c12; color: #f39c12; }}
.btn-chengyu {{ border-color: #9b59b6; color: #9b59b6; }}
.btn-changzhong {{ border-color: #1abc9c; color: #1abc9c; }}

.zoom-indicator {{
    position: absolute; top: 70px; left: 20px; z-index: 1000;
    background: rgba(52, 152, 219, 0.9); color: white; padding: 8px 15px;
    border-radius: 6px; font-size: 12px; font-weight: bold;
    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}}

.stats-panel {{
    position: absolute; top: 20px; right: 20px; z-index: 1000; width: 340px;
    background: rgba(255, 255, 255, 0.95); border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    font-family: "Microsoft YaHei", sans-serif; padding: 12px;
    max-height: calc(100vh - 40px);
    display: flex;
    flex-direction: column;
    transition: all 0.3s ease;
}}
.stats-panel.minimized {{
    width: auto;
    min-width: 200px;
}}
.stats-panel.minimized #stats-content {{
    display: none !important;
}}
.stats-title {{ font-size: 13px; color: #666; }}
.stats-total {{ font-size: 22px; color: #3498db; font-weight: bold; border-bottom: 2px solid #eee; padding-bottom: 8px; margin-bottom: 10px; }}
.stats-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 10px;
}}
.panel-controls {{
    display: flex;
    gap: 6px;
}}
.control-btn {{
    cursor: pointer;
    padding: 4px 8px;
    background: #f0f0f0;
    border-radius: 4px;
    font-size: 11px;
    color: #666;
    border: 1px solid #ddd;
    transition: all 0.2s;
}}
.control-btn:hover {{
    background: #e0e0e0;
    color: #333;
}}
.minimize-btn {{
    font-size: 14px;
    font-weight: bold;
    padding: 2px 8px;
}}
.stats-list {{ 
    flex: 1;
    overflow-y: auto; 
    font-size: 12px;
    min-height: 0;
}}

/* 三层层次结构样式 */
.stats-item {{ border-bottom: 1px solid #f0f0f0; padding: 5px 0; }}

/* 第1层: 资产类型 */
.level-1-header {{ 
    display: flex; justify-content: space-between; cursor: pointer; 
    font-weight: bold; padding: 6px 8px; background: #f8f9fa; 
    border-left: 3px solid #3498db; margin: 2px 0;
}}
.level-1-header:hover {{ background: #e3f2fd; }}

/* 第2层: REITs名称 */
.level-2-container {{ display: none; padding-left: 15px; }}
.level-2-item {{ margin: 4px 0; }}
.level-2-header {{ 
    cursor: pointer; color: #2c3e50; font-size: 11px; 
    font-weight: 600; padding: 4px 6px; background: #fff;
    border-left: 2px solid #95a5a6; display: flex; 
    justify-content: space-between; align-items: center;
}}
.level-2-header:hover {{ background: #ecf0f1; color: #3498db; }}

/* 第3层: 城市+底层资产 */
.level-3-detail {{ 
    display: none; padding: 6px 0 6px 15px; 
    color: #34495e; font-size: 10px; line-height: 1.8;
}}
.city-section {{ 
    margin: 4px 0; 
    padding: 4px 0;
    border-bottom: 1px dashed #ecf0f1;
}}
.city-section:last-child {{
    border-bottom: none;
}}
.city-name {{ 
    font-weight: 600; 
    color: #16a085; 
    margin-bottom: 2px;
}}
.asset-item {{ 
    padding: 1px 0 1px 12px;
    color: #7f8c8d;
    border-left: 1px solid #d5dbdb;
    margin-left: 4px;
}}

.search-box-container {{ position: absolute; top: 110px; left: 20px; z-index: 1000; width: 280px; }}
.search-input-wrapper {{ position: relative; display: flex; align-items: center; }}
#search-input {{ width: 100%; padding: 10px; border: 2px solid #3498db; border-radius: 6px; outline: none; box-sizing: border-box; font-size: 13px; }}
.search-results-list {{ background: white; border: 1px solid #ddd; max-height: 400px; overflow-y: auto; display: none; border-radius: 4px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); position: relative; margin-top: 4px; }}
.search-item {{ padding: 8px 12px; cursor: pointer; border-bottom: 1px solid #eee; transition: background 0.2s; }}
.search-item:hover {{ background: #f5f5f5; }}
.search-item:last-child {{ border-bottom: none; }}
.search-close-btn {{ position: absolute; top: 4px; right: 6px; cursor: pointer; color: #999; font-size: 16px; padding: 2px 6px; z-index: 10; }}
.search-close-btn:hover {{ color: #d32f2f; }}
.hint-text {{ position: absolute; top: 165px; left: 20px; z-index: 999; font-size: 11px; color: #888; max-width: 280px; line-height: 1.4; }}
</style>

<div class="nav-container">
    <button class="nav-btn btn-china" onclick="navigateTo(108, 35, 1.65)">全国</button>
    <button class="nav-btn btn-jingjinji" onclick="navigateTo(116.4, 39.5, 6)">京津冀</button>
    <button class="nav-btn btn-changsan" onclick="navigateTo(120.5, 31.5, 6)">长三角</button>
    <button class="nav-btn btn-wanqu" onclick="navigateTo(113.5, 22.8, 7)">大湾区</button>
    <button class="nav-btn btn-chengyu" onclick="navigateTo(105.5, 30.0, 6)">成渝</button>
    <button class="nav-btn btn-changzhong" onclick="navigateTo(114.3, 30.6, 6)">长江中游</button>
</div>

<div class="zoom-indicator" id="zoom-indicator">📍 当前:省份视图</div>

<div class="stats-panel" id="stats-panel">
    <div class="stats-header">
        <div>
            <div class="stats-title">上市 REITs 总数</div>
            <div class="stats-total">{total_reits_count} <span style="font-size:12px; font-weight:normal">只</span></div>
        </div>
        <div class="panel-controls">
            <button class="control-btn" id="collapse-all-btn" onclick="collapseAll()" title="收起全部展开项">收起全部</button>
            <button class="control-btn minimize-btn" id="minimize-btn" onclick="minimizePanel()" title="最小化面板">−</button>
        </div>
    </div>
    <div id="stats-content" class="stats-list"></div>
</div>

<div class="search-box-container">
    <div class="search-input-wrapper">
        <input type="text" id="search-input" placeholder="🔍 搜索REITs名称或底层资产" oninput="onSearchInput()" onkeydown="handleSearchKeydown(event)">
    </div>
    <div id="search-results" class="search-results-list"></div>
</div>
<div class="hint-text">💡 提示: 双击地图圆点可放大</div>

<script>
var provinceData = {province_data_json};
var cityData = {city_data_json};
var allSecurities = {securities_data_json};
var allCoords = {all_coords_json};
var visualMapColors = {visual_map_colors_json};
var assetHierarchy = {asset_hierarchy_json};
var maxProvinceCount = {max_province_count};
var maxCityCount = {max_city_count};
var currentZoom = 1.65;
var currentCenter = [108, 35];
var currentViewMode = 'province';
var provinceZoomThreshold = 3.5;
var initialProvinceCenter = [108, 35];
var initialProvinceZoom = 1.65;
var zoomSyncFrame = null;

function cloneOption(option) {{
    return option ? JSON.parse(JSON.stringify(option)) : null;
}}

function normalizeGeoOption(option) {{
    if (!option) return null;
    var normalized = cloneOption(option);
    if (normalized.geo && !Array.isArray(normalized.geo)) {{
        normalized.geo = [normalized.geo];
    }}
    return normalized;
}}

function getGeoState(option) {{
    if (!option) return {{}};
    return Array.isArray(option.geo) ? (option.geo[0] || {{}}) : (option.geo || {{}});
}}

function updateZoomIndicator() {{
    var indicator = document.getElementById('zoom-indicator');
    if (!indicator) return;
    indicator.innerHTML = currentViewMode === 'province' ?
        '📍 当前:省份视图' : '🏙️ 当前:城市视图';
}}

function getChartInstance() {{
    var chartDom = document.querySelector('.chart-container');
    return chartDom ? echarts.getInstanceByDom(chartDom) : null;
}}

function buildVisualPieces(maxCount) {{
    var pieces = [];
    for (var i = 1; i <= Math.min(10, maxCount); i++) {{
        pieces.push({{min: i, max: i, label: String(i)}});
    }}
    if (maxCount > 10) {{
        pieces.push({{min: 11, max: maxCount, label: "11+"}});
    }}
    return pieces;
}}

function buildFooterGraphic() {{
    return [{{
        type: 'group',
        left: 'center',
        bottom: 10,
        children: [{{
            type: 'text',
            left: 'center',
            bottom: 10,
            style: {{
                text: '数据来源: Wind, iFind | 原创制作: Qimai | 制作日期: {DISPLAY_DATE}',
                font: '12px Arial',
                fill: '#666666'
            }}
        }}]
    }}];
}}

function buildBaseGeo(center, zoom) {{
    return {{
        type: 'map',
        map: 'china',
        roam: true,
        center: center.slice(),
        zoom: zoom,
        itemStyle: {{
            areaColor: '#d0e1f9',
            borderColor: '#ffffff'
        }},
        label: {{
            show: false,
            color: '#1f2d3d',
            fontSize: 12
        }},
        emphasis: {{
            itemStyle: {{
                areaColor: '#b8d4f1'
            }},
            label: {{
                show: false,
                color: '#1f2d3d',
                fontSize: 12
            }}
        }}
    }};
}}

function buildProvinceOption(center, zoom) {{
    return {{
    animation: false,
    geo: [buildBaseGeo(center, zoom)],
    series: [{{
        name: '中国REITs分布',
        type: 'effectScatter',
        coordinateSystem: 'geo',
        data: provinceData.map(function(x) {{
            return {{
                name: x.name,
                value: [x.value[0], x.value[1], x.value[2]],
                tooltipHtml: x.value[3]
            }};
        }}),
        symbolSize: 13,
        showEffectOn: 'render',
        rippleEffect: {{
            brushType: 'stroke'
        }},
        label: {{
            show: true,
            formatter: '{{b}}',
            position: 'right',
            fontSize: 12,
            color: '#1f2d3d',
            opacity: 1,
            fontWeight: 600,
            textBorderColor: '#ffffff',
            textBorderWidth: 2
        }},
        emphasis: {{
            label: {{
                show: true,
                formatter: '{{b}}',
                position: 'right',
                fontSize: 12,
                color: '#1f2d3d',
                opacity: 1,
                fontWeight: 600,
                textBorderColor: '#ffffff',
                textBorderWidth: 2
            }}
        }},
        itemStyle: {{
            shadowBlur: 10,
            shadowColor: 'rgba(0, 0, 0, 0.3)',
            opacity: 1
        }}
    }}],
    tooltip: {{
        trigger: 'item',
        formatter: function(params) {{
            if (params.data && params.data.tooltipHtml) {{
                return params.data.tooltipHtml;
            }}
            return params.name || '';
        }}
    }},
    legend: {{
        show: true,
        top: 10,
        left: 'center',
        data: ['中国REITs分布']
    }},
    visualMap: {{
        piecewise: true,
        left: 10,
        bottom: 10,
        min: 1,
        max: maxProvinceCount,
        calculable: false,
        pieces: buildVisualPieces(maxProvinceCount),
        inRange: {{ color: visualMapColors }},
        textStyle: {{ fontSize: 12 }}
    }},
    graphic: buildFooterGraphic()
    }};
}}

function buildCityOption(center, zoom) {{
    var formatterJs = function(params) {{
        if (params.data && params.data.tooltipHtml) {{
            return params.data.tooltipHtml;
        }}
        return params.name || '';
    }};

    return {{
        animation: false,
        geo: [buildBaseGeo(center, zoom)],
        series: [{{
            name: '城市REITs分布',
            type: 'effectScatter',
            coordinateSystem: 'geo',
            data: cityData.map(function(x) {{
                return {{
                    name: x.name,
                    value: [x.value[0], x.value[1], x.value[2]],
                    tooltipHtml: x.value[3]
                }};
            }}),
            symbolSize: 12,
            showEffectOn: 'render',
            rippleEffect: {{
                brushType: 'stroke'
            }},
            label: {{
                show: false
            }},
            itemStyle: {{
                shadowBlur: 10,
                shadowColor: 'rgba(0, 0, 0, 0.3)',
                opacity: 1
            }}
        }}, {{
            name: '',
            type: 'custom',
            coordinateSystem: 'geo',
            silent: true,
            tooltip: {{
                show: false
            }},
            z: 30,
            data: cityData.map(function(x) {{ return [x.value[0], x.value[1], x.name]; }}),
            renderItem: function(params, api) {{
                var coord = api.coord([api.value(0), api.value(1)]);
                return {{
                    type: 'text',
                    x: coord[0] + 8,
                    y: coord[1] - 8,
                    style: {{
                        text: api.value(2),
                        fill: '#1f2d3d',
                        opacity: 1,
                        font: '600 12px Microsoft YaHei'
                    }}
                }};
            }}
        }}],
        tooltip: {{
            trigger: 'item',
            formatter: formatterJs
        }},
        legend: {{
            show: true,
            top: 10,
            left: 'center',
            data: ['城市REITs分布']
        }},
        visualMap: {{
            min: 1,
            max: maxCityCount,
            left: 10,
            bottom: 10,
            calculable: false,
            piecewise: true,
            pieces: buildVisualPieces(maxCityCount),
            inRange: {{
                color: visualMapColors
            }},
            textStyle: {{
                fontSize: 12
            }}
        }},
        graphic: buildFooterGraphic()
    }};
}}

function renderMapByMode(mode, center, zoom) {{
    var chart = getChartInstance();
    if (!chart) return;

    currentViewMode = mode;
    currentCenter = center.slice();
    currentZoom = zoom;

    var option = mode === 'province'
        ? buildProvinceOption(currentCenter, currentZoom)
        : buildCityOption(currentCenter, currentZoom);

    if (!option) return;

    chart.clear();
    chart.setOption(option, true);
    updateZoomIndicator();
}}

function updateMapData(zoom, forcedCenter) {{
    var chart = getChartInstance();
    if (!chart) return;

    var currentOption = chart.getOption() || {{}};
    var geoState = getGeoState(currentOption);
    var nextCenter = forcedCenter || geoState.center || currentCenter || initialProvinceCenter;
    var nextZoom = zoom || geoState.zoom || currentZoom || initialProvinceZoom;
    var nextMode = nextZoom >= provinceZoomThreshold ? 'city' : 'province';

    if (nextMode === 'province') {{
        nextCenter = initialProvinceCenter.slice();
        nextZoom = initialProvinceZoom;
    }}

    renderMapByMode(nextMode, nextCenter, nextZoom);
}}

function navigateTo(lon, lat, zoom) {{
    updateMapData(zoom, [lon, lat]);
}}

function renderStats() {{
    var container = document.getElementById('stats-content');
    container.innerHTML = '';
    
    // 遍历已排序的数据结构
    assetHierarchy.forEach(assetTypeData => {{
        var assetType = assetTypeData.asset_type;
        var reits = assetTypeData.reits;
        var reitsCount = assetTypeData.reits_count;
        
        // 计算总资产数
        var totalAssets = 0;
        reits.forEach(reit => {{
            reit.cities.forEach(cityData => {{
                totalAssets += cityData.asset_count;
            }});
        }});
        
        var level1Html = `
            <div class="stats-item">
                <div class="level-1-header" onclick="toggleNext(this, event)">
                    <span>${{assetType}}</span>
                    <span style="color:#3498db">${{reitsCount}}只 | ${{totalAssets}}项资产 ▾</span>
                </div>
                <div class="level-2-container">`;
        
        // 第2层: REITs (已按城市数排序)
        reits.forEach(reitData => {{
            var reit = reitData.reit;
            var cities = reitData.cities;
            var cityCount = reitData.city_count;
            
            var reitsAssetCount = 0;
            cities.forEach(cityData => {{
                reitsAssetCount += cityData.asset_count;
            }});
            
            level1Html += `
                <div class="level-2-item">
                    <div class="level-2-header" onclick="toggleNext(this, event)">
                        <span>📊 ${{reit}}</span>
                        <span style="color:#27ae60; font-size:10px">${{cityCount}}城 | ${{reitsAssetCount}}资产 ▾</span>
                    </div>
                    <div class="level-3-detail">`;
            
            // 第3层: 城市+底层资产 (已按资产数排序)
            cities.forEach(cityData => {{
                var city = cityData.city;
                var assets = cityData.assets;
                
                level1Html += `<div class="city-section">`;
                level1Html += `<div class="city-name">📍 ${{city}} (${{assets.length}}项)</div>`;
                
                if (assets.length > 0) {{
                    assets.forEach(asset => {{
                        level1Html += `<div class="asset-item">• ${{asset}}</div>`;
                    }});
                }} else {{
                    level1Html += `<div class="asset-item" style="color:#bdc3c7">暂无具体资产信息</div>`;
                }}
                
                level1Html += `</div>`;
            }});
            
            level1Html += `
                    </div>
                </div>`;
        }});
        
        level1Html += `
                </div>
            </div>`;
        
        container.insertAdjacentHTML('beforeend', level1Html);
    }});
}}

function toggleNext(el, evt) {{
    var target = el.nextElementSibling;
    if (target) {{
        target.style.display = (target.style.display === 'block') ? 'none' : 'block';
    }}
    if (evt) {{
        evt.stopPropagation();
    }}
}}

// ======== 统计面板控制功能 ========

// 最小化/展开面板
function minimizePanel() {{
    var panel = document.getElementById('stats-panel');
    var btn = document.getElementById('minimize-btn');
    
    if (panel.classList.contains('minimized')) {{
        // 展开面板
        panel.classList.remove('minimized');
        btn.textContent = '−';
        btn.title = '最小化面板';
    }} else {{
        // 最小化面板
        panel.classList.add('minimized');
        btn.textContent = '+';
        btn.title = '展开面板';
    }}
}}

// 收起全部展开的项目
function collapseAll() {{
    var content = document.getElementById('stats-content');
    
    // 收起所有第2层
    var allLevel2 = content.querySelectorAll('.level-2-container');
    allLevel2.forEach(function(el) {{
        el.style.display = 'none';
    }});
    
    // 收起所有第3层
    var allLevel3 = content.querySelectorAll('.level-3-detail');
    allLevel3.forEach(function(el) {{
        el.style.display = 'none';
    }});
}}

// ======== 优化后的搜索功能 ========

// 高亮关键词的辅助函数
function highlightKeyword(text, keyword) {{
    if (!keyword || !text) return text;
    var regex = new RegExp('(' + keyword.replace(/[.*+?^${{}}()|[\\]\\\\]/g, '\\\\$&') + ')', 'gi');
    return text.replace(regex, '<span style="background-color:#ffeb3b; color:#000; font-weight:bold; padding:1px 2px; border-radius:2px;">$1</span>');
}}

function onSearchInput() {{
    var kw = document.getElementById('search-input').value.trim();
    var resultDiv = document.getElementById('search-results');
    resultDiv.innerHTML = '';
    if (kw === '') {{ 
        resultDiv.style.display = 'none'; 
        clearSearchHighlight();
        return; 
    }}
    
    performSearch(kw);
}}

function handleSearchKeydown(event) {{
    if (event.key === 'Enter') {{
        event.preventDefault();
        var kw = document.getElementById('search-input').value.trim();
        if (kw !== '') {{
            performSearch(kw);
        }}
    }}
}}

function performSearch(kw) {{
    var resultDiv = document.getElementById('search-results');
    resultDiv.innerHTML = '';
    
    if (!kw || kw.trim() === '') {{
        resultDiv.style.display = 'none';
        clearSearchHighlight();
        return;
    }}
    
    var kwLower = kw.toLowerCase();
    var reitsResults = {{}};  // {{reit: {{assetType, cities: [{{city, assets, coord}}], hasMatch}}}}
    
    // 遍历所有资产数据，按REITs聚合
    assetHierarchy.forEach(assetTypeData => {{
        var assetType = assetTypeData.asset_type;
        assetTypeData.reits.forEach(reitData => {{
            var reit = reitData.reit;
            var reitMatches = reit.toLowerCase().indexOf(kwLower) !== -1;
            var hasAnyMatch = reitMatches;
            
            var citiesData = [];
            
            reitData.cities.forEach(cityData => {{
                var city = cityData.city;
                var allAssets = cityData.assets;
                
                // 检查是否有资产匹配
                var assetsMatch = allAssets.some(asset => 
                    asset.toLowerCase().indexOf(kwLower) !== -1
                );
                
                if (reitMatches || assetsMatch) {{
                    hasAnyMatch = true;
                    var cityCoord = allCoords[city] || [0, 0];
                    citiesData.push({{
                        city: city,
                        assets: allAssets,
                        lon: cityCoord[0],
                        lat: cityCoord[1]
                    }});
                }}
            }});
            
            // 只要该REITs有任何匹配，就加入结果
            if (hasAnyMatch && citiesData.length > 0) {{
                reitsResults[reit] = {{
                    reit: reit,
                    assetType: assetType,
                    cities: citiesData,
                    reitMatches: reitMatches
                }};
            }}
        }});
    }});
    
    // 显示搜索结果
    var resultsArray = Object.values(reitsResults);
    if (resultsArray.length > 0) {{
        resultDiv.style.display = 'block';
        
        // 添加关闭按钮
        var closeBtn = document.createElement('div');
        closeBtn.className = 'search-close-btn';
        closeBtn.innerHTML = '✕';
        closeBtn.title = '关闭';
        closeBtn.onclick = function(e) {{
            e.stopPropagation();
            resultDiv.style.display = 'none';
            resultDiv.dataset.pinned = 'false';
        }};
        resultDiv.appendChild(closeBtn);
        
        var headerDiv = document.createElement('div');
        headerDiv.style.cssText = 'padding:6px 12px; background:#e3f2fd; font-size:11px; font-weight:bold; color:#1976d2; border-bottom:1px solid #90caf9;';
        headerDiv.textContent = '搜索结果 (' + resultsArray.length + ' 个REITs)';
        resultDiv.appendChild(headerDiv);
        
        resultsArray.forEach(item => {{
            var div = document.createElement('div');
            div.className = 'search-item';
            div.style.cssText = 'padding:10px 12px; cursor:pointer; border-bottom:1px solid #eee;';
            
            // 计算总资产数
            var totalAssets = 0;
            item.cities.forEach(c => totalAssets += c.assets.length);
            
            // 构建城市和资产的HTML - 高亮关键词
            var citiesHtml = item.cities.map(cityData => {{
                var cityName = highlightKeyword(cityData.city, kw);
                var assetsHtml = cityData.assets.map(asset => {{
                    var highlightedAsset = highlightKeyword(asset, kw);
                    return '<div style="padding-left:12px; color:#666; font-size:10px; line-height:1.6;">• ' + highlightedAsset + '</div>';
                }}).join('');
                
                return `
                    <div style="margin-top:6px; padding:4px 0; border-left:2px solid #81c784; padding-left:8px;">
                        <div style="font-size:11px; color:#00796b; font-weight:600;">📍 ${{cityName}} (${{cityData.assets.length}}项资产)</div>
                        ${{assetsHtml}}
                    </div>
                `;
            }}).join('');
            
            // REITs标题 - 高亮关键词
            var reitTitle = highlightKeyword(item.reit, kw);
            var assetTypeText = highlightKeyword(item.assetType, kw);
            
            div.innerHTML = `
                <div style="font-weight:bold; color:#1976d2; font-size:12px; margin-bottom:4px;">
                    ${{reitTitle}} <span style="color:#666; font-size:10px;">(${{assetTypeText}})</span>
                </div>
                <div style="font-size:10px; color:#999; margin-bottom:4px;">
                    ${{item.cities.length}} 个城市 | ${{totalAssets}} 项底层资产
                </div>
                ${{citiesHtml}}
            `;
            
            // 点击任一城市定位到该城市
            div.onclick = function(e) {{
                // 默认定位到第一个城市
                var firstCity = item.cities[0];
                navigateTo(firstCity.lon, firstCity.lat, 10);
                keepSearchResultsVisible(item);
                
                // 高亮选中项
                var allItems = resultDiv.querySelectorAll('.search-item');
                allItems.forEach(i => i.style.background = '');
                div.style.background = '#e3f2fd';
            }};
            
            resultDiv.appendChild(div);
        }});
    }} else {{ 
        resultDiv.style.display = 'none'; 
    }}
}}

// 保持搜索结果可见，直到下一次搜索
function keepSearchResultsVisible(selectedItem) {{
    var resultDiv = document.getElementById('search-results');
    var searchInput = document.getElementById('search-input');
    
    // 标记当前为"固定显示"模式
    resultDiv.dataset.pinned = 'true';
    resultDiv.dataset.selectedReit = selectedItem.reit;
}}

// 清除搜索高亮
function clearSearchHighlight() {{
    var resultDiv = document.getElementById('search-results');
    resultDiv.dataset.pinned = 'false';
}}

document.addEventListener('click', e => {{
    var resultDiv = document.getElementById('search-results');
    var searchInput = document.getElementById('search-input');
    
    // 如果点击的不是搜索框，且不在固定模式下，则关闭搜索结果
    if (e.target.id !== 'search-input' && !resultDiv.contains(e.target)) {{
        if (resultDiv.dataset.pinned !== 'true') {{
            resultDiv.style.display = 'none';
        }}
    }}
}});

// 改进的滚轮控制 - 全屏任何地方都能缩放
window.addEventListener('wheel', function(e) {{
    var searchResults = document.getElementById('search-results');
    var statsContent = document.getElementById('stats-content');
    
    // 只在搜索结果和统计面板内部允许滚动
    if (searchResults && searchResults.contains(e.target)) {{
        // 检查是否到达边界
        if ((e.deltaY < 0 && searchResults.scrollTop === 0) || 
            (e.deltaY > 0 && searchResults.scrollTop + searchResults.clientHeight >= searchResults.scrollHeight)) {{
            e.preventDefault();
        }}
        return;
    }}
    
    if (statsContent && statsContent.contains(e.target)) {{
        // 检查是否到达边界
        if ((e.deltaY < 0 && statsContent.scrollTop === 0) || 
            (e.deltaY > 0 && statsContent.scrollTop + statsContent.clientHeight >= statsContent.scrollHeight)) {{
            e.preventDefault();
        }}
        return;
    }}
    
    // 其他所有区域都阻止默认滚动,让ECharts处理地图缩放
    e.preventDefault();
}}, {{ passive: false }});

window.addEventListener('load', function() {{
    renderStats();
    
    var chart = getChartInstance();
    if (chart) {{
        var initOption = normalizeGeoOption(chart.getOption());
        var initGeoState = getGeoState(initOption);
        if (initOption && initOption.geo && initOption.geo[0]) {{
            initialProvinceCenter = (initGeoState.center || initialProvinceCenter).slice();
            initialProvinceZoom = initGeoState.zoom || initialProvinceZoom;
            currentCenter = initialProvinceCenter.slice();
            currentZoom = initialProvinceZoom;
            currentViewMode = 'province';
            renderMapByMode('province', initialProvinceCenter.slice(), initialProvinceZoom);
        }}
        chart.on('georoam', function(params) {{
            if (params.zoom != null) {{
                if (zoomSyncFrame) {{
                    cancelAnimationFrame(zoomSyncFrame);
                }}
                zoomSyncFrame = requestAnimationFrame(function() {{
                    var option = chart.getOption();
                    var geoState = getGeoState(option);
                    var nextZoom = geoState.zoom || currentZoom;
                    var nextCenter = geoState.center || currentCenter;
                    currentCenter = nextCenter.slice ? nextCenter.slice() : nextCenter;
                    currentZoom = nextZoom;
                    var nextMode = nextZoom >= provinceZoomThreshold ? 'city' : 'province';
                    if (nextMode !== currentViewMode) {{
                        if (nextMode === 'province') {{
                            renderMapByMode('province', initialProvinceCenter.slice(), initialProvinceZoom);
                        }} else {{
                            renderMapByMode(nextMode, currentCenter, currentZoom);
                        }}
                    }} else {{
                        updateZoomIndicator();
                    }}
                }});
            }}
        }});
        
        chart.on('dblclick', function(params) {{
            if (params.componentType === 'series') {{
                var targetData = currentViewMode === 'province' ? provinceData : cityData;
                var location = targetData.find(c => c.name === params.name);
                if (location) {{
                    navigateTo(location.value[0], location.value[1], 8);
                }}
            }}
        }});
        
        window.addEventListener('resize', function() {{
            chart.resize();
        }});
    }}
}});
</script>
"""

html_content = re.sub(r"<body[^>]*>", lambda m: m.group(0) + "\n" + navigation_and_search_html, html_content, count=1)
html_content = html_content.replace(
    "<head>",
    """<head>
<title>C-REITs | China Real Estate Investment Trusts Map</title>
<link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' rx='18' fill='%230f172a'/%3E%3Crect x='28' y='20' width='44' height='60' fill='white'/%3E%3Crect x='38' y='30' width='8' height='8' fill='%230f172a'/%3E%3Crect x='54' y='30' width='8' height='8' fill='%230f172a'/%3E%3Crect x='38' y='46' width='8' height='8' fill='%230f172a'/%3E%3Crect x='54' y='46' width='8' height='8' fill='%230f172a'/%3E%3C/svg%3E">
"""
)
html_content = re.sub(
    r"""(chart_[A-Za-z0-9_]+\.setOption\(option_[A-Za-z0-9_]+\);)\s*window\.addEventListener\('resize',\s*function\(\)\s*\{\s*chart_[A-Za-z0-9_]+\.resize\(\);\s*\}\)\s*;?""",
    r"\1",
    html_content,
    count=1,
)

html_content = embed_local_js_assets(html_content)

with open(FINAL_HTML_PATH, "w", encoding="utf-8") as f:
    f.write(html_content)

if TEMP_HTML_PATH.exists():
    TEMP_HTML_PATH.unlink()

print(f"地图已生成:{FINAL_HTML_PATH.name}")

# %%

