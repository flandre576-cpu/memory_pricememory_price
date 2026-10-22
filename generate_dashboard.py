import os
import glob
import json
import pandas as pd
import numpy as np

def load_latest_data():
    search_pattern = os.path.join(".", "内存价格每日追踪_*.xlsx")
    matched_files = [f for f in glob.glob(search_pattern) if not os.path.basename(f).startswith('~')]
    if not matched_files:
        if os.path.exists("000000_历史基底数据.xlsx"):
            matched_files = ["000000_历史基底数据.xlsx"]
        else:
            raise FileNotFoundError("未找到任何 Excel 数据文件。")
    latest_excel = max(matched_files)
    
    df_dram = pd.read_excel(latest_excel, sheet_name='DRAM Spot Price', header=[0, 1], engine='openpyxl')
    df_flash = pd.read_excel(latest_excel, sheet_name='NAND Flash', header=[0, 1], engine='openpyxl')
    return df_dram, df_flash, os.path.basename(latest_excel)

def preprocess_df(df):
    # 展平多级表头并提取日期
    df.columns = [f"{col[0]}^{col[1]}" if 'Unnamed' not in str(col[0]) else col[1] for col in df.columns]
    date_col = [col for col in df.columns if '日期' in col][0]
    
    # 格式化日期为字符串 YYYY-MM-DD
    df[date_col] = pd.to_datetime(df[date_col]).dt.strftime('%Y-%m-%d')
    df = df.sort_values(date_col)
    df.set_index(date_col, inplace=True)
    
    # 🎯 核心修复：清理历史数据中重复的日期行，强制去重并保留当天最后/最新的一条
    df = df[~df.index.duplicated(keep='last')]
    
    return df

def generate_html():
    df_dram_raw, df_flash_raw, filename = load_latest_data()
    df_dram = preprocess_df(df_dram_raw)
    df_flash = preprocess_df(df_flash_raw)

    # ========================================================
    # 🔍 数据提取配置区
    # 当前逻辑：提取所有包含“盘平均”的列。
    # 如果想提取其他指标，只需将 '盘平均' 修改为 '日高点'、'日低点' 或 '盘高点'。
    # ========================================================
    dram_cols = [col for col in df_dram.columns if '盘平均' in col]
    flash_cols = [col for col in df_flash.columns if '盘平均' in col]

    # 合并 DRAM 和 Flash 数据
    df_combined = pd.concat([df_dram[dram_cols], df_flash[flash_cols]], axis=1)
    
    # 处理空值为 None (以适配 JSON)，并将所有数值四舍五入保留 2 位小数
    df_combined = df_combined.round(2).replace({np.nan: None})

    # 转化为 JS 变量所需的格式
    dates_list = df_combined.index.tolist()
    series_dict = df_combined.to_dict(orient='list')

    # 生成内嵌了数据的 HTML
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>TrendForce 存储半导体价格看板</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts/dist/echarts.min.js"></script>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background-color: #09090b; color: #fff; margin: 0; padding: 40px; }}
        .header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 40px; }}
        .title h1 {{ margin: 0; font-size: 28px; font-weight: 600; }}
        .title p {{ margin: 8px 0 0 0; color: #a1a1aa; font-size: 14px; }}
        .download-btn {{ background-color: #3b82f6; color: #fff; padding: 10px 20px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 14px; transition: background-color 0.2s; }}
        .download-btn:hover {{ background-color: #60a5fa; }}
        #charts-container {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(350px, 1fr)); gap: 24px; }}
        .chart-card {{ background-color: #121214; border: 1px solid #27272a; border-radius: 12px; padding: 20px; display: flex; flex-direction: column; transition: border-color 0.3s; }}
        .chart-card:hover {{ border-color: #3f3f46; }}
        .chart-header {{ display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 10px; }}
        .chart-title {{ font-size: 16px; font-weight: 600; color: #fff; }}
        .chart-latest-value {{ font-size: 24px; font-weight: 700; color: #3b82f6; font-family: monospace; }}
        .chart-body {{ height: 200px; min-height: 200px; width: 100%; }}
    </style>
</head>
<body>
    <div class="header">
        <div class="title">
            <h1>TrendForce 存储半导体价格看板</h1>
            <p>每日 DRAM 与 NAND Flash 现货均价追踪 (当前追踪表格: {filename})</p>
        </div>
        <a href="{filename}" class="download-btn" download>📥 下载原数据表</a>
    </div>
    <div id="charts-container"></div>

    <script>
        // 动态注入的数据
        const dates = {json.dumps(dates_list)};
        const seriesData = {json.dumps(series_dict)};
        
        const headers = Object.keys(seriesData);
        const container = document.getElementById('charts-container');

        headers.forEach((header, index) => {{
            const dataArray = seriesData[header];
            let latestValue = '--';
            // 倒序寻找最新有数据的天数
            for(let i = dataArray.length - 1; i >= 0; i--) {{
                if(dataArray[i] !== null && dataArray[i] !== undefined) {{
                    latestValue = Number(dataArray[i]).toFixed(2);
                    break;
                }}
            }}

            const card = document.createElement('div');
            card.className = 'chart-card';
            const chartId = 'chart-container-' + index;
            
            // 将 "DDR4 8Gx8^盘平均" 拆分展示
            const titleParts = header.split('^');
            const mainTitle = titleParts[0].trim();
            const unitTitle = titleParts.length > 1 ? titleParts[1].trim() : '';

            card.innerHTML = `
                <div class="chart-header">
                    <div class="chart-title">
                        ${{mainTitle}}
                        <div style="font-size: 11px; color: #71717a; font-weight: normal; margin-top: 2px;">${{unitTitle}}</div>
                    </div>
                    <div class="chart-latest-value">$${{latestValue}}</div>
                </div>
                <div id="${{chartId}}" class="chart-body"></div>
            `;
            container.appendChild(card);

            setTimeout(() => {{
                const chartDom = document.getElementById(chartId);
                if (!chartDom) return;
                
                const chart = echarts.init(chartDom);
                chart.setOption({{
                    tooltip: {{ 
                        trigger: 'axis',
                        formatter: function(params) {{
                            let val = params[0].value !== null ? '$' + Number(params[0].value).toFixed(2) : '无数据';
                            return `${{params[0].axisValue}}<br/>${{params[0].marker}} <b>${{val}}</b>`;
                        }}
                    }},
                    grid: {{ left: '2%', right: '2%', top: '5%', bottom: '5%', containLabel: true }},
                    xAxis: {{ 
                        type: 'category', 
                        data: dates, 
                        boundaryGap: false,
                        axisLine: {{ lineStyle: {{ color: '#3f3f46' }} }},
                        axisLabel: {{ color: '#71717a', fontSize: 10, margin: 12 }},
                        splitLine: {{ show: false }}
                    }},
                    yAxis: {{ 
                        type: 'value', 
                        scale: true, 
                        splitLine: {{ lineStyle: {{ color: '#27272a', type: 'dashed' }} }},
                        axisLabel: {{ color: '#71717a', fontSize: 10 }}
                    }},
                    series: [{{
                        data: dataArray,
                        type: 'line',
                        smooth: true,
                        connectNulls: true, 
                        symbol: 'circle',
                        symbolSize: 5,
                        showSymbol: true,   
                        lineStyle: {{ width: 2, color: '#3b82f6' }},
                        itemStyle: {{ color: '#3b82f6' }},
                        areaStyle: {{
                            color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                                {{ offset: 0, color: 'rgba(59, 130, 246, 0.2)' }},
                                {{ offset: 1, color: 'rgba(59, 130, 246, 0)' }}
                            ])
                        }}
                    }}]
                }});
                
                window.addEventListener('resize', () => chart.resize());
            }}, 50);
        }});
    </script>
</body>
</html>"""

    # 导出静态网页
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    print("✅ 静态看板 index.html 生成成功！")

if __name__ == "__main__":
    generate_html()
