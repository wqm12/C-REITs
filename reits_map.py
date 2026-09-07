import argparse
import json
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent
RUN_DATE_SUFFIX = datetime.now().strftime("%Y%m%d")
DATA_FILE = PROJECT_DIR / "REITs_location_updating.xlsx"
SOURCE_SCRIPT = PROJECT_DIR / "map_source_template.py"
PROCESS_DIR = PROJECT_DIR / "process_files" / f"process_{RUN_DATE_SUFFIX}_map_generation"
FINAL_HTML_PATH = PROJECT_DIR / f"REITs详细分布图_{RUN_DATE_SUFFIX}.html"
HTML_ARCHIVE_DIR = PROJECT_DIR / "历史版本" / "REIT地图历史版本" / "自动归档"

INFRA_CLASS = "基础设施REITs"
COMMERCIAL_CLASS = "商业不动产REITs"
BIG_CLASS_COL = "REITs大类"

ASSET_RENAME_MAP = {
    "园区": "产业园",
    "交通": "交通设施",
    "消费": "消费商业",
    "保障性租赁住房": "保租房",
    "仓储物流": "物流仓储",
    "新型": "数据中心",
    "市政": "市政公用",
    "水利": "水利",
}


def set_output_date(value):
    global RUN_DATE_SUFFIX, PROCESS_DIR, FINAL_HTML_PATH
    RUN_DATE_SUFFIX = datetime.strptime(value, "%Y-%m-%d").strftime("%Y%m%d")
    PROCESS_DIR = PROJECT_DIR / "process_files" / f"process_{RUN_DATE_SUFFIX}_map_generation"
    FINAL_HTML_PATH = PROJECT_DIR / f"REITs详细分布图_{RUN_DATE_SUFFIX}.html"


def replace_once(text, old, new):
    if old not in text:
        raise ValueError(f"未找到待替换片段: {old}")
    return text.replace(old, new, 1)


def archive_existing_html_results():
    HTML_ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for html_path in PROJECT_DIR.glob("REITs详细分布图_*.html"):
        if html_path == FINAL_HTML_PATH:
            archive_path = HTML_ARCHIVE_DIR / f"{html_path.stem}_before_{timestamp}.html"
            shutil.copy2(html_path, archive_path)
            continue

        archive_path = HTML_ARCHIVE_DIR / html_path.name
        if archive_path.exists():
            archive_path = HTML_ARCHIVE_DIR / f"{html_path.stem}_{timestamp}.html"
        shutil.copy2(html_path, archive_path)


def build_source_script(sheet_name, output_suffix):
    source_text = SOURCE_SCRIPT.read_text(encoding="utf-8")
    source_text = replace_once(source_text,
        'RUN_DATE_SUFFIX = datetime.now().strftime("%Y%m%d")',
        f'RUN_DATE_SUFFIX = "{RUN_DATE_SUFFIX}"')
    source_text = replace_once(source_text,
        'DISPLAY_DATE = datetime.now().strftime("%Y-%m")',
        f'DISPLAY_DATE = "{RUN_DATE_SUFFIX[:4]}-{RUN_DATE_SUFFIX[4:6]}"')
    source_text = replace_once(
        source_text,
        "df = pd.read_excel(DATA_FILE)",
        f'df = pd.read_excel(DATA_FILE, sheet_name="{sheet_name}")',
    )
    source_text = replace_once(
        source_text,
        'FINAL_HTML_PATH = PROJECT_DIR / f"REITs详细分布图_{RUN_DATE_SUFFIX}.html"',
        (
            'FINAL_HTML_PATH = PROJECT_DIR / "process_files" / '
            'f"process_{RUN_DATE_SUFFIX}_map_generation" / '
            f'f"REITs详细分布图_{{RUN_DATE_SUFFIX}}_{output_suffix}.html"'
        ),
    )
    source_path = PROCESS_DIR / f"source_{output_suffix}.py"
    source_path.write_text(source_text, encoding="utf-8")
    return source_path


def run_source_script(source_path):
    result = subprocess.run(
        [sys.executable, str(source_path)],
        cwd=str(PROJECT_DIR),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"{source_path.name} 执行失败\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        )


def extract_js_value(html, variable_name, next_variable_name):
    pattern = rf"var {re.escape(variable_name)} = (.*?);\s*var {re.escape(next_variable_name)} = "
    match = re.search(pattern, html, re.S)
    if not match:
        raise ValueError(f"未能提取 JS 变量: {variable_name}")
    return match.group(1).strip()


def extract_total(html):
    match = re.search(r'<div class="stats-total">(\d+)\s*<span', html)
    if not match:
        raise ValueError("未能提取 REITs 总数")
    return int(match.group(1))


def split_place_names(text):
    if pd.isna(text):
        return []
    return [part.strip() for part in re.split(r"[，,]", str(text)) if part.strip()]


def split_braced_groups(text):
    if pd.isna(text):
        return []
    value = str(text).strip()
    groups = []
    current = []
    in_group = False
    for char in value:
        if char == "{":
            current = []
            in_group = True
        elif char == "}":
            if in_group:
                groups.append("".join(current).strip())
            current = []
            in_group = False
        elif in_group:
            current.append(char)
    return groups or ([value] if value else [])


def split_group_items(group_text):
    normalized = str(group_text).replace(",", "，").replace("、", "，")
    return [part.strip() for part in normalized.split("，") if part.strip()]


def parse_grouped_items(text):
    return [split_group_items(group) for group in split_braced_groups(text)]


def normalize_asset_type(value):
    text = "" if pd.isna(value) else str(value).strip()
    text = text.replace("基础设施", "")
    return ASSET_RENAME_MAP.get(text, text)


def normalize_big_class(row):
    if BIG_CLASS_COL in row and not pd.isna(row[BIG_CLASS_COL]):
        text = str(row[BIG_CLASS_COL]).strip()
        if text:
            return text
    return COMMERCIAL_CLASS if str(row.get("资产类型", "")).strip() == "商业不动产" else INFRA_CLASS


def city_asset_rows(row):
    cities = split_place_names(row["项目城市"])
    asset_groups = parse_grouped_items(row["项目分类"])

    rows = []
    for index, city in enumerate(cities):
        assets = asset_groups[index] if index < len(asset_groups) else []
        if not assets:
            assets = [row["项目名称"]]
        rows.append((city, assets))
    return rows


def build_stats_hierarchy(df):
    hierarchy = {INFRA_CLASS: {}, COMMERCIAL_CLASS: {}}

    for _, row in df.iterrows():
        big_class = normalize_big_class(row)
        security = row["证券名称"]
        asset_type = normalize_asset_type(row["资产类型"])
        for city, assets in city_asset_rows(row):
            for asset in assets:
                add_asset(hierarchy, big_class, asset_type, security, city, asset)

    return sort_stats_hierarchy(hierarchy)


def add_asset(hierarchy, big_class, category, security, city, asset):
    hierarchy.setdefault(big_class, {})
    hierarchy[big_class].setdefault(category, {})
    hierarchy[big_class][category].setdefault(security, {})
    hierarchy[big_class][category][security].setdefault(city, [])
    if asset not in hierarchy[big_class][category][security][city]:
        hierarchy[big_class][category][security][city].append(asset)


def sort_stats_hierarchy(hierarchy):
    result = []
    for big_class in [INFRA_CLASS, COMMERCIAL_CLASS]:
        categories = []
        total_reits = set()
        total_assets = 0

        for category, reit_map in hierarchy.get(big_class, {}).items():
            reits_list = []
            category_reits = set()
            category_assets = 0

            for reit, city_map in reit_map.items():
                total_reits.add(reit)
                category_reits.add(reit)
                cities_list = []
                for city, assets in city_map.items():
                    category_assets += len(assets)
                    total_assets += len(assets)
                    cities_list.append({
                        "city": city,
                        "assets": assets,
                        "asset_count": len(assets),
                    })
                cities_list.sort(key=lambda item: item["asset_count"], reverse=True)
                reits_list.append({
                    "reit": reit,
                    "cities": cities_list,
                    "city_count": len(cities_list),
                })

            reits_list.sort(key=lambda item: item["city_count"], reverse=True)
            categories.append({
                "asset_type": category,
                "reits": reits_list,
                "total_cities": sum(reit["city_count"] for reit in reits_list),
                "reits_count": len(category_reits),
                "asset_count": category_assets,
            })

        categories.sort(key=lambda item: item["reits_count"], reverse=True)
        result.append({
            "big_class": big_class,
            "categories": categories,
            "reits_count": len(total_reits),
            "asset_count": total_assets,
        })
    return result


def extract_dataset(html, sheet_name, label, stats_title):
    df = pd.read_excel(DATA_FILE, sheet_name=sheet_name)
    return {
        "label": label,
        "stats_title": stats_title,
        "total": extract_total(html),
        "provinceData": extract_js_value(html, "provinceData", "cityData"),
        "cityData": extract_js_value(html, "cityData", "allSecurities"),
        "allSecurities": extract_js_value(html, "allSecurities", "allCoords"),
        "allCoords": extract_js_value(html, "allCoords", "visualMapColors"),
        "visualMapColors": extract_js_value(html, "visualMapColors", "assetHierarchy"),
        "assetHierarchy": json.dumps(build_stats_hierarchy(df), ensure_ascii=False),
        "maxProvinceCount": extract_js_value(html, "maxProvinceCount", "maxCityCount"),
        "maxCityCount": extract_js_value(html, "maxCityCount", "currentZoom"),
    }


def dataset_to_js(mode_key, dataset):
    return f"""
    {mode_key}: {{
        label: "{dataset["label"]}",
        statsTitle: "{dataset["stats_title"]}",
        total: {dataset["total"]},
        provinceData: {dataset["provinceData"]},
        cityData: {dataset["cityData"]},
        allSecurities: {dataset["allSecurities"]},
        allCoords: {dataset["allCoords"]},
        visualMapColors: {dataset["visualMapColors"]},
        assetHierarchy: {dataset["assetHierarchy"]},
        maxProvinceCount: {dataset["maxProvinceCount"]},
        maxCityCount: {dataset["maxCityCount"]}
    }}"""


def build_dataset_block(listed_html, waiting_html):
    listed = extract_dataset(listed_html, "已上市REITs", "已上市 REITs", "已上市 REITs 总数")
    waiting = extract_dataset(waiting_html, "未上市REITs", "未上市 REITs", "未上市 REITs 总数")
    return f"""window.reitsModeDatasets = {{
{dataset_to_js("listed", listed)},
{dataset_to_js("waiting", waiting)}
}};
var currentDatasetMode = 'listed';
var currentDataset = window.reitsModeDatasets[currentDatasetMode];
var provinceData = currentDataset.provinceData;
var cityData = currentDataset.cityData;
var allSecurities = currentDataset.allSecurities;
var allCoords = currentDataset.allCoords;
var visualMapColors = currentDataset.visualMapColors;
var assetHierarchy = currentDataset.assetHierarchy;
var maxProvinceCount = currentDataset.maxProvinceCount;
var maxCityCount = currentDataset.maxCityCount;

function applyReitsDatasetGlobals(dataset) {{
    provinceData = dataset.provinceData;
    cityData = dataset.cityData;
    allSecurities = dataset.allSecurities;
    allCoords = dataset.allCoords;
    visualMapColors = dataset.visualMapColors;
    assetHierarchy = dataset.assetHierarchy;
    maxProvinceCount = dataset.maxProvinceCount;
    maxCityCount = dataset.maxCityCount;
}}

function updateModeToggleState() {{
    document.querySelectorAll('.mode-toggle-btn').forEach(function(btn) {{
        btn.classList.toggle('active', btn.getAttribute('data-mode') === currentDatasetMode);
    }});
}}

function updateStatsHeader() {{
    var dataset = window.reitsModeDatasets[currentDatasetMode];
    var title = document.querySelector('.stats-title');
    var total = document.querySelector('.stats-total');
    if (title) title.textContent = dataset.statsTitle;
    if (total) total.innerHTML = dataset.total + ' <span style="font-size:12px; font-weight:normal">只</span>';
}}

function resetSearchState() {{
    var input = document.getElementById('search-input');
    var results = document.getElementById('search-results');
    if (input) input.value = '';
    if (results) {{
        results.innerHTML = '';
        results.style.display = 'none';
    }}
    if (typeof clearSearchHighlight === 'function') clearSearchHighlight();
}}

function switchReitsMode(mode) {{
    if (!window.reitsModeDatasets[mode] || mode === currentDatasetMode) return;
    currentDatasetMode = mode;
    currentDataset = window.reitsModeDatasets[currentDatasetMode];
    applyReitsDatasetGlobals(currentDataset);
    updateModeToggleState();
    updateStatsHeader();
    resetSearchState();
    renderStats();
    renderMapByMode('province', initialProvinceCenter.slice(), initialProvinceZoom);
}}"""


def add_mode_switch_ui(html):
    mode_css = """
.mode-toggle-container {
    position: absolute; top: 64px; left: 20px; z-index: 1000;
    display: flex; overflow: hidden; border: 2px solid #3498db;
    border-radius: 6px; background: rgba(255,255,255,0.95);
    box-shadow: 0 2px 8px rgba(0,0,0,0.12);
}
.mode-toggle-btn {
    border: 0; background: transparent; color: #3498db;
    padding: 7px 12px; cursor: pointer; font-weight: bold;
    font-size: 13px; font-family: "Microsoft YaHei", sans-serif;
}
.mode-toggle-btn + .mode-toggle-btn { border-left: 1px solid rgba(52,152,219,0.25); }
.mode-toggle-btn.active { background: #3498db; color: #fff; }
.mode-toggle-btn:hover { background: rgba(52,152,219,0.12); }
.mode-toggle-btn.active:hover { background: #2f86c1; }
.big-class-section { margin-bottom: 14px; }
.big-class-title {
    color: #315f7d; font-size: 15px; line-height: 1.2;
    font-weight: 700; margin: 8px 0 7px; padding: 8px 10px;
    background: linear-gradient(90deg, #e9f3f8 0%, #f7fafc 100%);
    border-bottom: 1px solid #dce9f0;
}
.big-class-empty {
    border: 1px dashed #9fb9c9; min-height: 58px; margin: 8px 0 6px;
    display: flex; align-items: center; justify-content: center;
    color: #78909c; font-size: 12px; background: #f6fafc;
}
.level-category-container { display: block; }
"""
    html = replace_once(
        html,
        ".btn-changzhong { border-color: #1abc9c; color: #1abc9c; }",
        ".btn-changzhong { border-color: #1abc9c; color: #1abc9c; }\n" + mode_css,
    )
    html = html.replace("position: absolute; top: 70px; left: 20px; z-index: 1000;", "position: absolute; top: 110px; left: 20px; z-index: 1000;", 1)
    html = html.replace(".search-box-container { position: absolute; top: 110px;", ".search-box-container { position: absolute; top: 150px;", 1)
    html = html.replace(".hint-text { position: absolute; top: 165px;", ".hint-text { position: absolute; top: 205px;", 1)

    toggle_html = """
<div class="mode-toggle-container" id="mode-toggle-container">
    <button class="mode-toggle-btn active" type="button" data-mode="listed" onclick="switchReitsMode('listed')">已上市 REITs</button>
    <button class="mode-toggle-btn" type="button" data-mode="waiting" onclick="switchReitsMode('waiting')">未上市 REITs</button>
</div>
"""
    return replace_once(
        html,
        '</div>\n\n<div class="zoom-indicator" id="zoom-indicator">',
        f'</div>\n{toggle_html}\n<div class="zoom-indicator" id="zoom-indicator">',
    )


def replace_function(html, function_name, replacement):
    marker = f"function {function_name}("
    start = html.find(marker)
    if start == -1:
        raise ValueError(f"未能替换 JS 函数: {function_name}")
    brace_start = html.find("{", start)
    if brace_start == -1:
        raise ValueError(f"未能定位 JS 函数起始括号: {function_name}")

    depth = 0
    in_string = None
    escaped = False
    for index in range(brace_start, len(html)):
        char = html[index]
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == in_string:
                in_string = None
            continue
        if char in ("'", '"', "`"):
            in_string = char
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return html[:start] + replacement + html[index + 1:]
    raise ValueError(f"未能定位 JS 函数结束括号: {function_name}")


RENDER_STATS_JS = r"""function renderStats() {
    var container = document.getElementById('stats-content');
    container.innerHTML = '';

    assetHierarchy.forEach(bigClassData => {
        var bigClass = bigClassData.big_class;
        var categories = bigClassData.categories || [];
        var bigClassHtml = `
            <div class="big-class-section">
                <div class="big-class-title">${bigClass}</div>`;

        if (categories.length === 0) {
            bigClassHtml += `<div class="big-class-empty">暂无项目</div></div>`;
            container.insertAdjacentHTML('beforeend', bigClassHtml);
            return;
        }

        categories.forEach(categoryData => {
            var assetType = categoryData.asset_type;
            var reits = categoryData.reits || [];
            var reitsCount = categoryData.reits_count || 0;
            var totalAssets = categoryData.asset_count || 0;

            bigClassHtml += `
                <div class="stats-item">
                    <div class="level-1-header" onclick="toggleNext(this, event)">
                        <span>${assetType}</span>
                        <span style="color:#3498db">${reitsCount}只 | ${totalAssets}项资产 ▾</span>
                    </div>
                    <div class="level-2-container">`;

            reits.forEach(reitData => {
                var reit = reitData.reit;
                var cities = reitData.cities || [];
                var cityCount = reitData.city_count || 0;
                var reitsAssetCount = 0;
                cities.forEach(cityData => {
                    reitsAssetCount += cityData.asset_count;
                });

                bigClassHtml += `
                    <div class="level-2-item">
                        <div class="level-2-header" onclick="toggleNext(this, event)">
                            <span>📊 ${reit}</span>
                            <span style="color:#27ae60; font-size:10px">${cityCount}城 | ${reitsAssetCount}资产 ▾</span>
                        </div>
                        <div class="level-3-detail">`;

                cities.forEach(cityData => {
                    var city = cityData.city;
                    var assets = cityData.assets || [];
                    bigClassHtml += `<div class="city-section">`;
                    bigClassHtml += `<div class="city-name">📍 ${city} (${assets.length}项)</div>`;
                    if (assets.length > 0) {
                        assets.forEach(asset => {
                            bigClassHtml += `<div class="asset-item">• ${asset}</div>`;
                        });
                    } else {
                        bigClassHtml += `<div class="asset-item" style="color:#bdc3c7">暂无具体资产信息</div>`;
                    }
                    bigClassHtml += `</div>`;
                });

                bigClassHtml += `
                        </div>
                    </div>`;
            });

            bigClassHtml += `
                    </div>
                </div>`;
        });

        bigClassHtml += `</div>`;
        container.insertAdjacentHTML('beforeend', bigClassHtml);
    });
}"""


COLLAPSE_ALL_JS = r"""function collapseAll() {
    var content = document.getElementById('stats-content');
    var allLevel2 = content.querySelectorAll('.level-2-container');
    allLevel2.forEach(function(el) {
        el.style.display = 'none';
    });
    var allLevel3 = content.querySelectorAll('.level-3-detail');
    allLevel3.forEach(function(el) {
        el.style.display = 'none';
    });
}"""


PERFORM_SEARCH_JS = r"""function renderSearchItemDetails(cities, kw, expanded) {
    var cityLimit = expanded ? cities.length : 3;
    var html = '';

    cities.slice(0, cityLimit).forEach(cityData => {
        html += `<div class="search-city-row" style="font-size:10px; color:#34495e; margin:3px 0;">
            <span style="color:#16a085">📍 ${cityData.city}</span> (${cityData.assets.length}项资产)
        </div>`;

        var assetLimit = expanded ? cityData.assets.length : 3;
        cityData.assets.slice(0, assetLimit).forEach(asset => {
            html += `<div class="search-asset-row" style="font-size:10px; color:#7f8c8d; padding-left:15px; line-height:1.4;">
                • ${highlightKeyword(asset, kw)}
            </div>`;
        });

        if (!expanded && cityData.assets.length > assetLimit) {
            html += `<div data-search-expand="assets" style="font-size:10px; color:#5b9bd5; padding-left:15px; cursor:pointer;">
                还有 ${cityData.assets.length - assetLimit} 项资产...
            </div>`;
        }
    });

    if (!expanded && cities.length > cityLimit) {
        html += `<div data-search-expand="cities" style="font-size:10px; color:#5b9bd5; margin-top:4px; cursor:pointer;">
            还有 ${cities.length - cityLimit} 个城市...
        </div>`;
    }
    return html;
}

function expandCurrentReit(div, cities, kw, resultDiv) {
    var details = div.querySelector('.search-item-details');
    if (details) {
        details.innerHTML = renderSearchItemDetails(cities, kw, true);
        div.dataset.expanded = 'true';
    }
    if (cities.length > 0) {
        var firstCity = cities[0];
        navigateTo(firstCity.lon, firstCity.lat, 8);
        resultDiv.dataset.pinned = 'true';
    }
}

function performSearch(kw) {
    var resultDiv = document.getElementById('search-results');
    resultDiv.innerHTML = '';

    if (!kw || kw.trim() === '') {
        resultDiv.style.display = 'none';
        clearSearchHighlight();
        return;
    }

    var kwLower = kw.toLowerCase();
    var reitsResults = {};

    assetHierarchy.forEach(bigClassData => {
        (bigClassData.categories || []).forEach(categoryData => {
            var assetType = categoryData.asset_type;
            (categoryData.reits || []).forEach(reitData => {
                var reit = reitData.reit;
                var reitMatches = reit.toLowerCase().indexOf(kwLower) !== -1;
                var allCitiesData = [];
                var anyAssetMatches = false;

                (reitData.cities || []).forEach(cityData => {
                    var city = cityData.city;
                    var allAssets = cityData.assets || [];
                    var assetsMatch = allAssets.some(asset =>
                        asset.toLowerCase().indexOf(kwLower) !== -1
                    );
                    anyAssetMatches = anyAssetMatches || assetsMatch;
                    var cityCoord = allCoords[city] || [0, 0];
                    allCitiesData.push({
                        city: city,
                        assets: allAssets,
                        lon: cityCoord[0],
                        lat: cityCoord[1]
                    });
                });

                if (reitMatches || anyAssetMatches) {
                    if (!reitsResults[reit]) {
                        reitsResults[reit] = {
                            reit: reit,
                            assetType: assetType,
                            cities: [],
                            reitMatches: reitMatches
                        };
                    }
                    reitsResults[reit].cities = reitsResults[reit].cities.concat(allCitiesData);
                }
            });
        });
    });

    var resultsArray = Object.values(reitsResults);
    if (resultsArray.length > 0) {
        var showFullDetails = resultsArray.length === 1;
        resultDiv.style.display = 'block';

        var closeBtn = document.createElement('div');
        closeBtn.className = 'search-close-btn';
        closeBtn.innerHTML = '✕';
        closeBtn.title = '关闭';
        closeBtn.onclick = function(e) {
            e.stopPropagation();
            resultDiv.style.display = 'none';
            resultDiv.dataset.pinned = 'false';
        };
        resultDiv.appendChild(closeBtn);

        var headerDiv = document.createElement('div');
        headerDiv.style.cssText = 'padding:6px 12px; background:#e3f2fd; font-size:11px; font-weight:bold; color:#1976d2; border-bottom:1px solid #90caf9;';
        headerDiv.textContent = '搜索结果 (' + resultsArray.length + ' 个REITs)';
        resultDiv.appendChild(headerDiv);

        resultsArray.forEach(item => {
            var uniqueCities = {};
            item.cities.forEach(cityData => {
                if (!uniqueCities[cityData.city]) {
                    uniqueCities[cityData.city] = {
                        city: cityData.city,
                        assets: [],
                        lon: cityData.lon,
                        lat: cityData.lat
                    };
                }
                cityData.assets.forEach(asset => {
                    if (uniqueCities[cityData.city].assets.indexOf(asset) === -1) {
                        uniqueCities[cityData.city].assets.push(asset);
                    }
                });
            });
            var cities = Object.values(uniqueCities);
            var totalAssets = cities.reduce((sum, city) => sum + city.assets.length, 0);

            var div = document.createElement('div');
            div.className = 'search-item';
            div.style.cssText = 'padding:10px 12px; border-bottom:1px solid #eee; cursor:pointer;';

            var html = `<div style="font-weight:bold; color:#2c3e50; margin-bottom:4px;">
                ${highlightKeyword(item.reit, kw)} <span style="color:#7f8c8d; font-size:11px">(${item.assetType})</span>
            </div>`;
            html += `<div style="font-size:10px; color:#95a5a6; margin-bottom:6px;">
                ${cities.length} 个城市 | ${totalAssets} 项底层资产
            </div>`;
            html += `<div class="search-item-details">${renderSearchItemDetails(cities, kw, showFullDetails)}</div>`;

            div.innerHTML = html;
            div.onclick = function(e) {
                if (e.target.closest('[data-search-expand]')) {
                    e.stopPropagation();
                    expandCurrentReit(div, cities, kw, resultDiv);
                    return;
                }
                if (cities.length > 0) {
                    var firstCity = cities[0];
                    navigateTo(firstCity.lon, firstCity.lat, 8);
                    resultDiv.dataset.pinned = 'true';
                }
            };
            resultDiv.appendChild(div);
        });
    } else {
        resultDiv.style.display = 'block';
        resultDiv.innerHTML = '<div style="padding:12px; color:#999; text-align:center; font-size:12px;">未找到相关结果</div>';
    }
}"""


def merge_dual_mode_html(listed_html, waiting_html):
    dataset_block = build_dataset_block(listed_html, waiting_html)
    html = add_mode_switch_ui(listed_html)
    data_block_pattern = (
        r"var provinceData = .*?;\s*"
        r"var cityData = .*?;\s*"
        r"var allSecurities = .*?;\s*"
        r"var allCoords = .*?;\s*"
        r"var visualMapColors = .*?;\s*"
        r"var assetHierarchy = .*?;\s*"
        r"var maxProvinceCount = .*?;\s*"
        r"var maxCityCount = .*?;\s*"
    )
    html, count = re.subn(data_block_pattern, dataset_block + "\n", html, count=1, flags=re.S)
    if count != 1:
        raise ValueError("未能替换地图数据变量区块")

    html = replace_function(html, "renderStats", RENDER_STATS_JS)
    html = replace_function(html, "collapseAll", COLLAPSE_ALL_JS)
    html = replace_function(html, "performSearch", PERFORM_SEARCH_JS)
    html = html.replace(
        '<div class="stats-title">上市 REITs 总数</div>',
        '<div class="stats-title">已上市 REITs 总数</div>',
        1,
    )
    html = html.replace(
        "window.addEventListener('load', function() {\n    renderStats();",
        "window.addEventListener('load', function() {\n    updateStatsHeader();\n    updateModeToggleState();\n    renderStats();",
        1,
    )
    return html


def main():
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"缺少数据文件: {DATA_FILE}")
    if not SOURCE_SCRIPT.exists():
        raise FileNotFoundError(f"缺少历史源脚本: {SOURCE_SCRIPT}")

    PROCESS_DIR.mkdir(parents=True, exist_ok=True)
    archive_existing_html_results()
    listed_source = build_source_script("已上市REITs", "listed_source")
    waiting_source = build_source_script("未上市REITs", "waiting_source")

    run_source_script(listed_source)
    run_source_script(waiting_source)

    listed_html_path = PROCESS_DIR / f"REITs详细分布图_{RUN_DATE_SUFFIX}_listed_source.html"
    waiting_html_path = PROCESS_DIR / f"REITs详细分布图_{RUN_DATE_SUFFIX}_waiting_source.html"
    listed_html = listed_html_path.read_text(encoding="utf-8")
    waiting_html = waiting_html_path.read_text(encoding="utf-8")

    final_html = merge_dual_mode_html(listed_html, waiting_html)
    FINAL_HTML_PATH.write_text(final_html, encoding="utf-8")
    print(f"地图已生成:{FINAL_HTML_PATH.name}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="从定位表生成 REITs 地图")
    parser.add_argument("--date", help="清单命名日，格式 YYYY-MM-DD；默认当天")
    args = parser.parse_args()
    if args.date:
        set_output_date(args.date)
    main()
