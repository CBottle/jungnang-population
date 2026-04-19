# -*- coding: utf-8 -*-
"""중랑구 연령별 인구 대시보드 생성기"""

import pandas as pd
import numpy as np
import json
import re
import sys
from pathlib import Path

try:
    import requests
except ImportError:
    print("requests 설치 필요: pip install requests")
    sys.exit(1)

# ── 설정 ─────────────────────────────────────────────────────────────────────

FILES = {
    2022: 'data/서울특별시_중랑구_연령별인구현황_20220203.xlsx',
    2023: 'data/서울특별시_중랑구_연령별인구현황_20230427.xlsx',
    2025: 'data/서울특별시_중랑구_연령별_인구수_현황_20250430.xlsx',
    2026: 'data/서울특별시 중랑구 연령별 인구수 현황_20260130.xlsx',
}

YEARS = [2022, 2023, 2025, 2026]

AGE_GROUPS = {
    '전체':   range(0, 110),
    '0-9세':  range(0, 10),
    '10-19세': range(10, 20),
    '20-29세': range(20, 30),
    '30-39세': range(30, 40),
    '40-49세': range(40, 50),
    '50-59세': range(50, 60),
    '60-69세': range(60, 70),
    '70-79세': range(70, 80),
    '80-89세': range(80, 90),
    '90세+':  range(90, 110),
}

GEOJSON_URLS = [
    'https://raw.githubusercontent.com/vuski/admdongkor/master/ver20231001/HangJeongDong_ver20231001.geojson',
    'https://raw.githubusercontent.com/vuski/admdongkor/master/ver20220101/HangJeongDong_ver20220101.geojson',
]

# ── 데이터 파싱 ───────────────────────────────────────────────────────────────

def parse_excel(filepath, year):
    df_raw = pd.read_excel(filepath, sheet_name=0, header=None, dtype=object)

    # 헤더행 찾기 (중랑구 포함)
    header_row = None
    for i in range(len(df_raw)):
        if '중랑구' in [str(v) for v in df_raw.iloc[i].values]:
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"헤더를 찾을 수 없음: {filepath}")

    # 동 이름 → 컬럼 인덱스 매핑
    header = df_raw.iloc[header_row]
    dong_to_col = {}
    for ci, val in enumerate(header):
        s = str(val).strip()
        if s in ('NaN', 'nan', '', '구분', '인구수', '중랑구'):
            continue
        if '동' in s:
            dong_to_col[s] = ci

    # 나이 데이터 시작행 찾기 (0세)
    age_start = None
    for i in range(header_row + 1, len(df_raw)):
        if re.match(r'^\d+세$', str(df_raw.iloc[i, 1]).strip()):
            age_start = i
            break
    if age_start is None:
        raise ValueError(f"연령 데이터를 찾을 수 없음: {filepath}")

    records = []
    current_age = None
    for i in range(age_start, len(df_raw)):
        row = df_raw.iloc[i]
        col1 = str(row.iloc[1]).strip()
        col2 = str(row.iloc[2]).strip()

        m = re.match(r'^(\d+)세$', col1)
        if m:
            current_age = int(m.group(1))

        if current_age is None or col2 not in ('계', '남', '여'):
            continue

        for dong, ci in dong_to_col.items():
            v = row.iloc[ci]
            if pd.notna(v):
                try:
                    pop = int(float(str(v).replace(',', '')))
                    records.append({
                        'year': year, 'dong': dong,
                        'age': current_age, 'gender': col2,
                        'population': pop
                    })
                except (ValueError, TypeError):
                    pass

    return pd.DataFrame(records)


def load_all_data():
    dfs = []
    for year, fp in FILES.items():
        print(f"  {year}년 파싱 중...", end=' ')
        df = parse_excel(fp, year)
        print(f"{len(df):,}행, {df['dong'].nunique()}개 동")
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)


# ── 집계 ─────────────────────────────────────────────────────────────────────

def aggregate(df):
    """반환: {str(year): {gender: {age_group: {dong: pop}}}}"""
    dongs = sorted(df['dong'].unique().tolist())
    result = {}
    for year in YEARS:
        ydf = df[df['year'] == year]
        result[str(year)] = {}
        for gender in ('계', '남', '여'):
            gdf = ydf[ydf['gender'] == gender]
            result[str(year)][gender] = {}
            for ag, rng in AGE_GROUPS.items():
                adf = gdf[gdf['age'].isin(rng)]
                dong_pop = adf.groupby('dong')['population'].sum().to_dict()
                result[str(year)][gender][ag] = {d: int(dong_pop.get(d, 0)) for d in dongs}
    return result, dongs


def compute_changes(pop_data, dongs):
    """전년(이전 데이터) 대비 변화율. 첫 해는 null."""
    prev = {2023: 2022, 2025: 2023, 2026: 2025}
    result = {}
    for year in YEARS:
        sy = str(year)
        result[sy] = {}
        for gender in ('계', '남', '여'):
            result[sy][gender] = {}
            for ag in AGE_GROUPS:
                if year not in prev:
                    result[sy][gender][ag] = {d: None for d in dongs}
                else:
                    py = str(prev[year])
                    result[sy][gender][ag] = {}
                    for dong in dongs:
                        curr = pop_data[sy][gender][ag].get(dong, 0)
                        prv  = pop_data[py][gender][ag].get(dong, 0)
                        if prv > 0:
                            result[sy][gender][ag][dong] = round((curr - prv) / prv * 100, 2)
                        else:
                            result[sy][gender][ag][dong] = None
    return result


# ── GeoJSON ───────────────────────────────────────────────────────────────────

def load_geojson(dongs):
    full = None
    for url in GEOJSON_URLS:
        try:
            print(f"  GeoJSON 다운로드: {url.split('/')[-1]}")
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            full = r.json()
            break
        except Exception as e:
            print(f"  실패: {e}")

    if full is None:
        return None

    # 중랑구만 필터 (sgg == '11260' 또는 adm_cd 11070으로 시작)
    features = [f for f in full['features']
                if f['properties'].get('sgg') == '11260'
                or str(f['properties'].get('adm_cd', '')).startswith('11070')]
    print(f"  중랑구 동 {len(features)}개 추출")

    def _norm(s):
        # 제거: 제(第 관형사), 중점 → 마침표
        return s.replace('제', '').replace('·', '.').replace('ㆍ', '.').strip()

    # dong_name 정규화 (GeoJSON adm_nm ↔ Excel 동명 매핑)
    # Excel 동 정규화 사전 미리 생성
    norm_map = {_norm(d): d for d in dongs}

    for feat in features:
        full_nm = feat['properties'].get('adm_nm', '')
        short_nm = full_nm.split()[-1] if full_nm.strip() else full_nm
        matched = norm_map.get(_norm(short_nm))
        feat['properties']['dong_name'] = matched or short_nm
        feat['id'] = feat['properties']['dong_name']

    geo_dongs = {f['properties']['dong_name'] for f in features}
    missing = set(dongs) - geo_dongs
    if missing:
        print(f"  [!] GeoJSON 미매칭 동: {missing}")

    return {'type': 'FeatureCollection', 'features': features}


# ── HTML 템플릿 ───────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>중랑구 연령별 인구 대시보드</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js" charset="utf-8"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
*{margin:0;padding:0;box-sizing:border-box}

body{
  font-family:'Inter','Apple SD Gothic Neo','Malgun Gothic',sans-serif;
  height:100vh;display:flex;flex-direction:column;overflow:hidden;
  background: linear-gradient(135deg,#0f0c29 0%,#1a1040 35%,#0d1f3c 65%,#0a1628 100%);
  color:#fff;
}
body::before{
  content:'';position:fixed;inset:0;
  background: radial-gradient(ellipse at 20% 20%,rgba(120,80,255,.18) 0%,transparent 60%),
              radial-gradient(ellipse at 80% 80%,rgba(0,180,255,.12) 0%,transparent 60%);
  pointer-events:none;z-index:0;
}

/* 글라스 믹스인 */
.glass{
  background:rgba(255,255,255,.07);
  backdrop-filter:blur(24px) saturate(180%);
  -webkit-backdrop-filter:blur(24px) saturate(180%);
  border:1px solid rgba(255,255,255,.13);
  border-radius:18px;
}

header{
  background:rgba(255,255,255,.05);
  backdrop-filter:blur(30px);
  -webkit-backdrop-filter:blur(30px);
  border-bottom:1px solid rgba(255,255,255,.1);
  padding:14px 22px;display:flex;align-items:center;gap:14px;flex-shrink:0;
  position:relative;z-index:10;
}
header h1{font-size:1.15rem;font-weight:700;letter-spacing:-.3px}
header p{font-size:.75rem;opacity:.5;margin-top:2px;font-weight:300;letter-spacing:.3px}

.controls{
  background:rgba(255,255,255,.05);
  backdrop-filter:blur(20px);
  -webkit-backdrop-filter:blur(20px);
  border-bottom:1px solid rgba(255,255,255,.08);
  padding:10px 22px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;flex-shrink:0;
  position:relative;z-index:10;
}
.cg{display:flex;align-items:center;gap:7px}
.cg label{font-size:.72rem;font-weight:600;color:rgba(255,255,255,.55);white-space:nowrap;letter-spacing:.4px;text-transform:uppercase}
select{
  padding:6px 10px;
  border:1px solid rgba(255,255,255,.18);
  border-radius:10px;
  font-size:.78rem;
  background:rgba(255,255,255,.1);
  backdrop-filter:blur(10px);
  color:#fff;
  cursor:pointer;outline:none;font-family:inherit;
  transition:all .2s;
}
select:focus{border-color:rgba(120,160,255,.6);background:rgba(255,255,255,.15)}
select option{background:#1a1040;color:#fff}

.mode-toggle{
  display:flex;
  border:1px solid rgba(255,255,255,.18);
  border-radius:10px;overflow:hidden;
  background:rgba(255,255,255,.07);
}
.mbtn{
  padding:6px 14px;font-size:.75rem;cursor:pointer;border:none;
  background:transparent;color:rgba(255,255,255,.55);
  font-family:inherit;font-weight:500;
  transition:all .2s;letter-spacing:.2px;
}
.mbtn.active{
  background:rgba(120,160,255,.35);
  color:#fff;font-weight:700;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.2);
}
.hint-txt{font-size:.72rem;color:rgba(255,255,255,.3);letter-spacing:.2px}

.main{display:flex;flex:1;min-height:0;padding:10px;gap:10px;position:relative;z-index:5}

.map-wrap{
  flex:1;min-width:0;
  border-radius:18px;overflow:hidden;
  border:1px solid rgba(255,255,255,.12);
  box-shadow:0 8px 32px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.1);
}
#map{width:100%;height:100%}

.sidebar{width:360px;display:flex;flex-direction:column;gap:8px;overflow:hidden}
.card{
  background:rgba(255,255,255,.07);
  backdrop-filter:blur(24px) saturate(180%);
  -webkit-backdrop-filter:blur(24px) saturate(180%);
  border:1px solid rgba(255,255,255,.12);
  border-radius:18px;padding:13px 15px;
  box-shadow:0 4px 24px rgba(0,0,0,.25),inset 0 1px 0 rgba(255,255,255,.1);
  display:flex;flex-direction:column;
}
.card-title{
  font-size:.68rem;font-weight:700;
  color:rgba(255,255,255,.45);
  text-transform:uppercase;letter-spacing:.7px;margin-bottom:6px;
}
.chart-wrap{flex:1;min-height:0}
#card-stats{flex-shrink:0}
#card-bar{flex:1.1}
#card-line{flex:1}

.hint{font-size:.75rem;color:rgba(255,255,255,.25);text-align:center;padding:14px 0}
.stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:4px}
.stat-item{
  background:rgba(255,255,255,.07);
  border:1px solid rgba(255,255,255,.1);
  border-radius:12px;padding:8px 10px;
}
.stat-item .sl{font-size:.65rem;color:rgba(255,255,255,.4);margin-bottom:2px;letter-spacing:.3px}
.stat-item .sv{font-size:.95rem;font-weight:700;color:#fff}
.pos{color:#ff6b6b}.neg{color:#74b9ff}

.year-bar{
  background:rgba(255,255,255,.05);
  backdrop-filter:blur(20px);
  -webkit-backdrop-filter:blur(20px);
  border-top:1px solid rgba(255,255,255,.08);
  padding:9px 22px 11px;display:flex;align-items:center;gap:12px;flex-shrink:0;
  position:relative;z-index:10;
}
.yl{font-size:.72rem;font-weight:600;color:rgba(255,255,255,.45);white-space:nowrap;letter-spacing:.3px}
.yv{font-size:1.05rem;font-weight:700;color:#a78bfa;min-width:44px;text-align:center}
input[type=range]{
  flex:1;-webkit-appearance:none;height:3px;
  background:rgba(255,255,255,.15);border-radius:2px;outline:none;cursor:pointer;
}
input[type=range]::-webkit-slider-thumb{
  -webkit-appearance:none;width:16px;height:16px;border-radius:50%;
  background:linear-gradient(135deg,#a78bfa,#60a5fa);
  cursor:pointer;box-shadow:0 2px 8px rgba(167,139,250,.5);
}
.play-btn{
  width:30px;height:30px;border:none;
  background:linear-gradient(135deg,#7c3aed,#2563eb);
  color:#fff;border-radius:50%;cursor:pointer;font-size:10px;
  flex-shrink:0;display:flex;align-items:center;justify-content:center;
  box-shadow:0 4px 12px rgba(124,58,237,.4);
  transition:transform .15s,box-shadow .15s;
}
.play-btn:hover{transform:scale(1.08);box-shadow:0 6px 16px rgba(124,58,237,.5)}
</style>
</head>
<body>

<header>
  <div>
    <h1>서울 중랑구 연령별 인구 대시보드</h1>
    <p>2022 · 2023 · 2025 · 2026년 &nbsp;|&nbsp; 동별 · 연령대별 · 성별 인구 현황</p>
  </div>
</header>

<div class="controls">
  <div class="cg">
    <label>연령대</label>
    <select id="selAge">
      <option value="전체">전체</option>
      <option value="0-9세">0–9세</option>
      <option value="10-19세">10–19세</option>
      <option value="20-29세">20–29세</option>
      <option value="30-39세">30–39세</option>
      <option value="40-49세">40–49세</option>
      <option value="50-59세">50–59세</option>
      <option value="60-69세">60–69세</option>
      <option value="70-79세">70–79세</option>
      <option value="80-89세">80–89세</option>
      <option value="90세+">90세+</option>
    </select>
  </div>
  <div class="cg">
    <label>성별</label>
    <select id="selGender">
      <option value="계">전체</option>
      <option value="남">남성</option>
      <option value="여">여성</option>
    </select>
  </div>
  <div class="cg">
    <label>지도 표시</label>
    <div class="mode-toggle">
      <button class="mbtn active" id="btnPop" onclick="setMode('인구수')">인구수</button>
      <button class="mbtn" id="btnChg" onclick="setMode('변화율')">변화율</button>
    </div>
  </div>
  <div class="cg" style="margin-left:auto">
    <span class="hint-txt">* 지도를 클릭하면 동별 상세 분석</span>
  </div>
</div>

<div class="main">
  <div class="map-wrap"><div id="map"></div></div>
  <div class="sidebar">
    <div class="card" id="card-stats">
      <div class="card-title" id="stats-title">동 선택 시 표시</div>
      <div id="stats-body"><div class="hint">지도에서 동을 클릭하세요</div></div>
    </div>
    <div class="card" id="card-bar">
      <div class="card-title">연령대별 인구 분포 <span style="font-weight:400;text-transform:none" id="bar-sub"></span></div>
      <div class="chart-wrap" id="barChart"></div>
    </div>
    <div class="card" id="card-line">
      <div class="card-title">연도별 인구 추이 <span style="font-weight:400;text-transform:none" id="line-sub"></span></div>
      <div class="chart-wrap" id="lineChart"></div>
    </div>
  </div>
</div>

<div class="year-bar">
  <button class="play-btn" id="playBtn" onclick="togglePlay()">▶</button>
  <span class="yl">연도</span>
  <span class="yv" id="yearDisplay">2022</span>
  <input type="range" id="yearSlider" min="0" max="3" step="1" value="0" oninput="onSlider(this.value)">
  <span style="font-size:.72rem;color:#aaa">2022 · 2023 · 2025 · 2026</span>
</div>

<script>
// ── 데이터 ──────────────────────────────────────────────────────────────────
const YEARS  = [2022,2023,2025,2026];
const AGES   = ['전체','0-9세','10-19세','20-29세','30-39세','40-49세','50-59세','60-69세','70-79세','80-89세','90세+'];
const GEOJSON   = __GEOJSON__;
const POP       = __POP_DATA__;
const CHANGE    = __CHANGE_DATA__;
const DONGS     = __DONGS__;

// ── 상태 ──────────────────────────────────────────────────────────────────
let S = {yi:0, age:'전체', gender:'계', mode:'인구수', dong:null, playing:false, timer:null};

const Y = () => String(YEARS[S.yi]);

// ── 지도 ──────────────────────────────────────────────────────────────────
function mapZ() {
  const src = S.mode==='인구수' ? POP : CHANGE;
  return DONGS.map(d => { const v=src[Y()][S.gender][S.age][d]; return (v===null||v===undefined)?0:v; });
}

function initMap() {
  const z = mapZ();
  const isChg = S.mode==='변화율';
  Plotly.newPlot('map',[{
    type:'choroplethmapbox',
    geojson:GEOJSON,
    locations:DONGS,
    featureidkey:'properties.dong_name',
    z:z,
    colorscale: isChg ? [[0,'#2471a3'],[0.5,'#f5f5f5'],[1,'#c0392b']] : [[0,'#eaf4fb'],[0.4,'#5dade2'],[0.7,'#1a6fa8'],[1,'#0a2d4a']],
    zmid: isChg ? 0 : undefined,
    colorbar:{title:{text:isChg?'변화율(%)':'인구수'},thickness:12,len:.6,tickfont:{size:9}},
    hovertemplate:'<b>%{location}</b><br>'+(isChg?'%{z:.1f}%':'%{z:,}명')+'<extra></extra>',
    marker:{opacity:.9,line:{color:'white',width:1.5}},
  }],{
    mapbox:{style:'carto-darkmatter',center:{lat:37.606,lon:127.093},zoom:12.3},
    margin:{t:0,b:0,l:0,r:0},
    paper_bgcolor:'transparent',
  },{responsive:true,displayModeBar:false,scrollZoom:true});

  document.getElementById('map').on('plotly_click', e => {
    if(e.points.length>0) { S.dong=e.points[0].location; updateSidebar(); }
  });
}

function updateMap() {
  const z = mapZ();
  const isChg = S.mode==='변화율';
  Plotly.restyle('map',{
    z:[z],
    colorscale:[isChg?[[0,'#2471a3'],[0.5,'#f5f5f5'],[1,'#c0392b']]:[[0,'#eaf4fb'],[0.4,'#5dade2'],[0.7,'#1a6fa8'],[1,'#0a2d4a']]],
    zmid:[isChg?0:undefined],
    'colorbar.title.text':[isChg?'변화율(%)':'인구수'],
    hovertemplate:['<b>%{location}</b><br>'+(isChg?'%{z:.1f}%':'%{z:,}명')+'<extra></extra>'],
  });
}

// ── 사이드바 ───────────────────────────────────────────────────────────────
function updateSidebar() {
  if(!S.dong) return;
  const d=S.dong, y=Y();

  // 통계 카드
  document.getElementById('stats-title').textContent=d;
  const tot  = POP[y]['계']['전체'][d]||0;
  const male = POP[y]['남']['전체'][d]||0;
  const fem  = POP[y]['여']['전체'][d]||0;
  const chg  = CHANGE[y]['계']['전체'][d];
  const chgStr = chg===null||chg===undefined ? 'N/A' : (chg>0?'+':'')+chg.toFixed(1)+'%';
  const chgCls = chg===null||chg===undefined ? '' : (chg>0?'pos':'neg');
  document.getElementById('stats-body').innerHTML=`
    <div class="stat-grid">
      <div class="stat-item"><div class="sl">전체 인구</div><div class="sv">${tot.toLocaleString()}명</div></div>
      <div class="stat-item"><div class="sl">전년 대비</div><div class="sv ${chgCls}">${chgStr}</div></div>
      <div class="stat-item"><div class="sl">남성</div><div class="sv">${male.toLocaleString()}명</div></div>
      <div class="stat-item"><div class="sl">여성</div><div class="sv">${fem.toLocaleString()}명</div></div>
    </div>`;

  updateBar(d, y);
  updateLine(d);
}

function updateBar(dong, y) {
  document.getElementById('bar-sub').textContent=`— ${dong} · ${y}년`;
  const ags = AGES.filter(a=>a!=='전체');
  const maleV = ags.map(a=>-(POP[y]['남'][a][dong]||0));
  const femV  = ags.map(a=> (POP[y]['여'][a][dong]||0));

  const glassFont = {color:'rgba(255,255,255,.7)',family:'Inter,Apple SD Gothic Neo,sans-serif'};
  Plotly.react('barChart',[
    {type:'bar',name:'남성',y:ags,x:maleV,orientation:'h',marker:{color:'rgba(96,165,250,.8)',line:{width:0}}},
    {type:'bar',name:'여성',y:ags,x:femV,orientation:'h',marker:{color:'rgba(244,114,182,.8)',line:{width:0}}},
  ],{
    barmode:'overlay',
    margin:{t:4,b:28,l:62,r:4},
    xaxis:{
      tickfont:{size:8,...glassFont},tickformat:',d',
      title:{text:'← 남성   여성 →',font:{size:9,...glassFont}},
      zeroline:true,zerolinecolor:'rgba(255,255,255,.15)',zerolinewidth:1.5,
      gridcolor:'rgba(255,255,255,.06)',
    },
    yaxis:{tickfont:{size:9,...glassFont},autorange:'reversed',gridcolor:'rgba(255,255,255,.06)'},
    showlegend:false,
    paper_bgcolor:'transparent',plot_bgcolor:'transparent',
  },{responsive:true,displayModeBar:false});
}

function updateLine(dong) {
  document.getElementById('line-sub').textContent=`— ${dong}`;
  const palette=['#e74c3c','#e67e22','#f39c12','#2ecc71','#1abc9c','#3498db','#9b59b6','#34495e','#e91e63','#795548'];
  const ags = AGES.filter(a=>a!=='전체');

  const traces = [{
    type:'scatter',mode:'lines+markers',name:'전체',
    x:YEARS, y:YEARS.map(yr=>POP[String(yr)][S.gender]['전체'][dong]||0),
    line:{color:'#2c3e50',width:3},marker:{size:8},
    visible: S.age==='전체' ? true : 'legendonly',
  }];

  ags.forEach((ag,i)=>{
    traces.push({
      type:'scatter',mode:'lines+markers',name:ag,
      x:YEARS, y:YEARS.map(yr=>POP[String(yr)][S.gender][ag][dong]||0),
      line:{color:palette[i%palette.length],width:2},marker:{size:5},
      visible: ag===S.age ? true : 'legendonly',
    });
  });

  const gf = {color:'rgba(255,255,255,.7)',family:'Inter,Apple SD Gothic Neo,sans-serif'};
  Plotly.react('lineChart',traces,{
    margin:{t:4,b:28,l:48,r:4},
    xaxis:{tickvals:YEARS,ticktext:YEARS.map(String),tickfont:{size:8,...gf},gridcolor:'rgba(255,255,255,.06)'},
    yaxis:{tickfont:{size:8,...gf},title:{text:'인구수',font:{size:9,...gf}},gridcolor:'rgba(255,255,255,.06)'},
    legend:{font:{size:8,...gf},orientation:'h',y:-0.35,tracegroupgap:0},
    paper_bgcolor:'transparent',plot_bgcolor:'transparent',
    height:undefined,
  },{responsive:true,displayModeBar:false});
}

// ── 컨트롤 ─────────────────────────────────────────────────────────────────
document.getElementById('selAge').addEventListener('change',e=>{S.age=e.target.value;updateMap();if(S.dong)updateSidebar();});
document.getElementById('selGender').addEventListener('change',e=>{S.gender=e.target.value;updateMap();if(S.dong)updateSidebar();});

function setMode(m){
  S.mode=m;
  document.getElementById('btnPop').classList.toggle('active',m==='인구수');
  document.getElementById('btnChg').classList.toggle('active',m==='변화율');
  updateMap();
}

function onSlider(v){
  S.yi=parseInt(v);
  document.getElementById('yearDisplay').textContent=YEARS[S.yi];
  updateMap();
  if(S.dong)updateSidebar();
}

function togglePlay(){
  if(S.playing){
    clearInterval(S.timer);S.playing=false;
    document.getElementById('playBtn').innerHTML='▶';
  } else {
    S.playing=true;
    document.getElementById('playBtn').innerHTML='⏸';
    S.timer=setInterval(()=>{
      S.yi=(S.yi+1)%YEARS.length;
      const sl=document.getElementById('yearSlider');
      sl.value=S.yi;
      document.getElementById('yearDisplay').textContent=YEARS[S.yi];
      updateMap();
      if(S.dong)updateSidebar();
    },1600);
  }
}

// ── 초기화 ─────────────────────────────────────────────────────────────────
initMap();
</script>
</body>
</html>
"""


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    print("=== 중랑구 인구 대시보드 생성기 ===\n")

    print("[1] 엑셀 데이터 파싱")
    df = load_all_data()
    pop_data, dongs = aggregate(df)
    print(f"  동 목록 ({len(dongs)}개): {dongs}")

    print("\n[2] 변화율 계산")
    change_data = compute_changes(pop_data, dongs)
    print("  완료")

    print("\n[3] GeoJSON 로드")
    geojson = load_geojson(dongs)
    if geojson is None:
        print("  ⚠ GeoJSON 없이 진행 불가. 네트워크 확인 후 재실행하세요.")
        return

    print("\n[4] HTML 생성")
    html = HTML
    html = html.replace('__GEOJSON__',  json.dumps(geojson,     ensure_ascii=False))
    html = html.replace('__POP_DATA__', json.dumps(pop_data,    ensure_ascii=False))
    html = html.replace('__CHANGE_DATA__', json.dumps(change_data, ensure_ascii=False))
    html = html.replace('__DONGS__',    json.dumps(dongs,        ensure_ascii=False))

    out = Path('dashboard.html')
    out.write_text(html, encoding='utf-8')
    size_kb = out.stat().st_size / 1024
    print(f"  저장 완료: {out.absolute()}")
    print(f"  파일 크기: {size_kb:.0f} KB")
    print("\n[완료] dashboard.html을 브라우저에서 열어보세요.")


if __name__ == '__main__':
    main()
