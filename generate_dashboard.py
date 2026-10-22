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
    return df_dram, df_flash

def preprocess_df(df):
    # 展平多级表头
    df.columns = [f"{col[0]}^{col[1]}" if 'Unnamed' not in str(col[0]) else col[1] for col in df.columns]
    date_col = [col for col in df.columns if '日期' in str(col)][0]
    
    # 日期格式化与排序
    df[date_col] = pd.to_datetime(df[date_col]).dt.strftime('%Y-%m-%d')
    df = df.sort_values(date_col)
    df.set_index(date_col, inplace=True)
    
    # 🎯 核心修复：强制去重，防止历史表中存在重复日期导致合并崩溃
    df = df[~df.index.duplicated(keep='last')]
    return df

def generate_html():
    df_dram_raw, df_flash_raw = load_latest_data()
    df_dram = preprocess_df(df_dram_raw)
    df_flash = preprocess_df(df_flash_raw)

    # 提取盘平均指标
    dram_cols = [col for col in df_dram.columns if str(col).endswith('盘平均')]
    flash_cols = [col for col in df_flash.columns if str(col).endswith('盘平均')]

    # 横向合并 DRAM 和 Flash 数据
    df_combined = pd.concat([df_dram[dram_cols], df_flash[flash_cols]], axis=1)
    
    # 🎯 核心修复：合并后强制按日期时间线重新升序排列，杜绝时间错乱乱序
    df_combined = df_combined.sort_index()
    
    # 处理空值，保留两位小数
    df_combined = df_combined.round(2).replace({np.nan: None})

    # 将数据转化为前端渲染所需的列表与字典
    dates_list = df_combined.index.tolist()
    series_dict = df_combined.to_dict(orient='list')

    # =========================================================================
    # 注入 ECharts 前端模板
    # =========================================================================
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>TrendForce 存储半导体价格看板</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts/dist/echarts.min.js"></script>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background-color: #09090b; color: #fff; margin: 0; padding: 40px; }}
        .header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 40px; }}
        
        /* 💎 国信专属品牌徽标样式（已调大字号与内边距） */
        .team-badge {{ display: inline-flex; align-items: center; background-color: rgba(103, 222, 249, 0.1); color: #67def9; padding: 10px 18px; border-radius: 24px; font-size: 16px; font-weight: 700; letter-spacing: 1.2px; margin-bottom: 16px; border: 1px solid rgba(103, 222, 249, 0.3); box-shadow: 0 0 12px rgba(103, 222, 249, 0.08); }}
        .team-badge svg {{ width: 18px; height: 18px; margin-right: 8px; fill: currentColor; }}
        
        .title h1 {{ margin: 0; font-size: 28px; font-weight: 600; }}
        .title p {{ margin: 8px 0 0 0; color: #a1a1aa; font-size: 14px; }}
        
        #charts-container {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(350px, 1fr)); gap: 24px; }}
        
        .chart-card {{ background-color: #121214; border: 1px solid #27272a; border-radius: 12px; padding: 20px; display: flex; flex-direction: column; transition: border-color 0.3s; }}
        .chart-card:hover {{ border-color: #3f3f46; }}
        .chart-header {{ display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 10px; }}
        .chart-title {{ font-size: 15px; font-weight: 600; color: #fff; line-height: 1.2; }}
        .chart-latest-value {{ font-size: 24px; font-weight: 700; color: #67def9; font-family: monospace; }}
        .chart-body {{ height: 200px; min-height: 200px; width: 100%; }}
    </style>
</head>
<body>
    <div class="header">
        <div class="title">
            <!-- 💎 放大后的国信徽标 -->
            <div class="team-badge">
                <svg viewBox="0 0 24 24"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/></svg>
                国信计算机&通信
            </div>
            <h1>TrendForce 存储半导体价格看板</h1>
            <p>每日 DRAM 与 NAND Flash 现货均价高频追踪</p>
        </div>
    </div>
    <div id="charts-container"></div>

    <script>
        // 动态注入的数据集
        const dates = {json.dumps(dates_list, ensure_ascii=False)};
        const seriesData = {json.dumps(series_dict, ensure_ascii=False)};
        
        const headers = Object.keys(seriesData);
        const container = document.getElementById('charts-container');

        headers.forEach((header, index) => {{
            const dataArray = seriesData[header];
            let latestValue = '--';
            
            for(let i = dataArray.length - 1; i >= 0; i--) {{
                if(dataArray[i] !== null && dataArray[i] !== undefined) {{
                    latestValue = Number(dataArray[i]).toFixed(2);
                    break;
                }}
            }}

            const card = document.createElement('div');
            card.className = 'chart-card';
            const chartId = 'chart-container-' + index;
            
            const titleParts = header.split('^');
            const mainTitle = titleParts[0].replace(' (SUSD)', '').replace(' ($USD)', '').trim();
            const unitTitle = titleParts.length > 1 ? titleParts[1].trim() + ' (USD)' : 'USD';

            card.innerHTML = `
                <div class="chart-header">
                    <div class="chart-title">
                        ${{mainTitle}}
                        <div style="font-size: 11px; color: #71717a; font-weight: normal; margin-top: 4px;">${{unitTitle}}</div>
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
                            let val = params[0].value !== null && params[0].value !== undefined ? '$' + Number(params[0].value).toFixed(2) : '无数据';
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
                        showSymbol: false,
                        lineStyle: {{ width: 2, color: '#67def9' }},
                        itemStyle: {{ color: '#67def9' }},
                        areaStyle: {{
                            color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                                {{ offset: 0, color: 'rgba(103, 222, 249, 0.2)' }},
                                {{ offset: 1, color: 'rgba(103, 222, 249, 0)' }}
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
    print("✅ ECharts 静态看板 index.html 生成成功！")

if __name__ == "__main__":
    generate_html()
