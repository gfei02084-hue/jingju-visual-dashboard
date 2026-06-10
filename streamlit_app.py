# -*- coding: utf-8 -*-
"""
数绘梨园｜京剧剧本多维可视分析平台 V5.0

运行前安装：
    python -m pip install streamlit pandas openpyxl plotly scikit-learn networkx

运行：
    python -m streamlit run jingju_dashboard_v5.py

说明：
    本程序读取你第一至第四问已经生成的 Excel/CSV 结果，不重新处理 PDF。
    左侧输入“前四问结果所在文件夹”，程序会自动递归查找相关结果文件。
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import html
import itertools
import re

import networkx as nx
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from sklearn.decomposition import PCA
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.preprocessing import StandardScaler


# =========================================================
# 0. 基础配置
# =========================================================

APP_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_ROOT = APP_DIR / "data"
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(
    page_title="数绘梨园｜京剧剧本多维可视分析平台",
    page_icon="🎭",
    layout="wide",
    initial_sidebar_state="expanded",
)

PAPER = "#F7F3EA"
INK = "#25252D"
RED = "#8C1D18"
GOLD = "#C89B3C"
BLUE = "#315B7D"
JADE = "#6C9A8B"

THEME_ORDER = [
    "婚恋姻缘", "礼教家庭", "忠义家国", "权力公案", "女性主体",
    "才学科举", "侠义救助", "悲剧生死", "喜剧讽刺", "神仙仪式",
]
BIG_ORDER = ["生", "旦", "净", "丑", "末", "其他", "未标注", "未知"]
PERIOD_ORDER = [
    "神话传说", "先秦两汉", "三国魏晋南北朝", "隋唐五代",
    "宋辽金元", "明清时期", "近现代", "民间传说/不明时代", "未知",
]

HANGDANG_COLORS = {
    "生": "#315B7D", "旦": "#D98032", "净": "#A52A2A", "丑": "#5D9B8C",
    "末": "#6B8E4E", "其他": "#8B6F9C", "未标注": "#B5ADA3",
    "未知": "#B5ADA3", "nan": "#B5ADA3",
}

RELATION_COLORS = {
    "亲属婚恋": "#D98062", "君臣权力": "#6A4C93", "冲突敌对": "#B92727",
    "救助协作": "#2A8F83", "信息传递": "#457B9D", "直接对白": "#D9A441",
    "称谓提及": "#7D9B76", "共场关系": "#9E9E9E", "普通互动": "#B0A79E",
}

MODE_COLORS = {
    "阵营对抗—忠义家国—军情出征型": "#315B7D",
    "权力中介—公案救助—反转平冤型": "#8C1D18",
    "亲密闭合—婚恋礼教—情感阻隔型": "#D98062",
    "仪式聚集—神仙祥瑞—圆满祝颂型": "#C89B3C",
    "多线复合—多主题—综合推进型": "#6C9A8B",
}

NETWORK_METRIC_LABELS = {
    "node_count": "角色数量",
    "edge_count": "关系边数",
    "density": "网络密度",
    "avg_degree": "平均度",
    "avg_clustering": "聚类系数",
    "modularity": "模块度",
}

MODE_EXPLAIN = {
    "阵营对抗—忠义家国—军情出征型":
        "君臣权力、阵营对抗和敌我冲突构成人物骨架，忠义家国与战争牺牲是主题核心，叙事通常沿军情—出征—交锋—决战—凯旋或牺牲推进。",
    "权力中介—公案救助—反转平冤型":
        "官府、权贵、受害者和救助者形成权力链条，权力公案与侠义救助是核心主题，剧情依靠对白、审判、周旋和反转推进。",
    "亲密闭合—婚恋礼教—情感阻隔型":
        "男女主人公、父母、丫鬟和媒介人物形成亲密网络，婚恋姻缘、礼教家庭与悲剧生死相互叠加，唱白交替与情绪积累推动剧情。",
    "仪式聚集—神仙祥瑞—圆满祝颂型":
        "人物关系以群体共场和等级秩序为主，冲突较弱，主题集中于神仙仪式、祥瑞祝寿和群体庆典，结构多为登场—献礼—祝颂—圆满。",
    "多线复合—多主题—综合推进型":
        "关系结构、主题组合和叙事方式均较为复合，多个角色群体与多条主题线索并行推进。",
}

FILE_CANDIDATES = {
    "role": [
        "问题一_角色行当特征表_含时期分析字段.xlsx",
        "问题一_角色行当特征表_含预测结果.xlsx",
        "问题一_角色行当特征表_预测前.xlsx",
    ],
    "network": [
        "第二问_剧目网络结构指标表.xlsx",
        "第二问_剧目网络指标表.xlsx",
    ],
    "edge": [
        "第二问_角色互动关系边表.xlsx",
        "第二问_角色关系边表.xlsx",
    ],
    "theme": [
        "第三问_主题强度与主题组合结果表.xlsx",
        "第三问_主题强度结果表.xlsx",
        "第三问_剧目主题聚类结果.xlsx",
    ],
    "narrative": [
        "第四问_剧本叙事结构指标表.xlsx",
        "第四问_剧本级叙事指标表.xlsx",
    ],
    "scene": [
        "第四问_场次叙事指标表_预览版.xlsx",
        "第四问_场次叙事指标表_含完整场次文本.csv",
        "第四问_场次叙事指标表.csv",
    ],
}

# =========================================================
# 1. 页面美化 CSS
# =========================================================

st.markdown(
    f"""
<style>
.stApp {{
  background:
    radial-gradient(circle at 92% 4%, rgba(200,155,60,0.08), transparent 18%),
    radial-gradient(circle at 8% 96%, rgba(140,29,24,0.06), transparent 20%),
    linear-gradient(180deg, #fffdf8 0%, #f7f3ea 100%);
}}
.block-container {{
  max-width: 1680px;
  padding-top: 0.8rem;
  padding-bottom: 1.3rem;
}}
[data-testid="stSidebar"] {{
  background: linear-gradient(180deg, #2b2a31 0%, #222126 100%);
  border-right: 1px solid rgba(200,155,60,0.35);
}}
[data-testid="stSidebar"] * {{
  color: #f7f3ea;
}}
.hero {{
  position: relative;
  overflow: hidden;
  border-radius: 18px;
  padding: 18px 24px 16px 24px;
  margin-bottom: 12px;
  color: #fffdf8;
  background:
    linear-gradient(110deg, rgba(37,37,45,0.98) 0%, rgba(61,30,32,0.97) 56%, rgba(140,29,24,0.93) 100%);
  border: 1px solid rgba(200,155,60,0.55);
  box-shadow: 0 8px 24px rgba(37,37,45,0.14);
}}
.hero::before {{
  content: "◈";
  position: absolute;
  right: 28px;
  top: -34px;
  font-size: 146px;
  color: rgba(200,155,60,0.13);
  transform: rotate(18deg);
}}
.hero::after {{
  content: "脸谱 · 水袖 · 锣鼓点 · 祥云";
  position: absolute;
  right: 28px;
  bottom: 12px;
  font-size: 12px;
  letter-spacing: 4px;
  color: rgba(247,243,234,0.56);
}}
.hero-title {{
  font-size: 32px;
  font-weight: 800;
  letter-spacing: 2px;
  margin: 0;
}}
.hero-sub {{
  margin-top: 5px;
  color: rgba(247,243,234,0.78);
  font-size: 14px;
}}
.gold-divider {{
  height: 1px;
  background: linear-gradient(90deg, transparent, rgba(200,155,60,0.95), transparent);
  margin: 6px 0 12px 0;
}}
.metric-card {{
  background: rgba(255,253,248,0.94);
  border: 1px solid rgba(200,155,60,0.42);
  border-top: 3px solid {RED};
  border-radius: 13px;
  min-height: 82px;
  padding: 10px 12px;
  box-shadow: 0 3px 12px rgba(37,37,45,0.05);
  overflow-wrap: anywhere;
}}
.metric-label {{
  font-size: 12px;
  color: #786e63;
  margin-bottom: 4px;
}}
.metric-value {{
  color: {INK};
  font-weight: 800;
  font-size: 21px;
  line-height: 1.23;
}}
.section-title {{
  display: flex;
  align-items: center;
  gap: 9px;
  color: {INK};
  font-size: 19px;
  font-weight: 800;
  margin: 10px 0 7px 0;
}}
.section-title::before {{
  content: "";
  width: 5px;
  height: 22px;
  border-radius: 4px;
  background: linear-gradient(180deg, {GOLD}, {RED});
}}
.mode-card {{
  border-radius: 13px;
  padding: 12px 13px;
  min-height: 140px;
  color: #fffdf8;
  box-shadow: 0 5px 15px rgba(37,37,45,0.10);
}}
.mode-card h4 {{
  margin: 0 0 6px 0;
  font-size: 15px;
}}
.mode-card p {{
  margin: 0;
  font-size: 12px;
  line-height: 1.55;
  color: rgba(255,253,248,0.84);
}}
.role-card {{
  border-radius: 16px;
  padding: 18px;
  min-height: 260px;
  background:
    radial-gradient(circle at 16% 22%, rgba(200,155,60,0.14), transparent 27%),
    linear-gradient(145deg, rgba(255,253,248,0.98), rgba(247,243,234,0.95));
  border: 1px solid rgba(140,29,24,0.22);
  box-shadow: inset 0 0 0 3px rgba(200,155,60,0.07), 0 5px 16px rgba(37,37,45,0.05);
}}
.badge {{
  display: inline-block;
  padding: 4px 10px;
  margin: 2px 4px 2px 0;
  border-radius: 999px;
  color: #fffdf8;
  background: {RED};
  font-size: 12px;
}}
.triangle {{
  position: relative;
  height: 330px;
  border-radius: 18px;
  background:
    radial-gradient(circle at center, rgba(200,155,60,0.10), transparent 44%),
    rgba(255,253,248,0.84);
  border: 1px solid rgba(200,155,60,0.30);
}}
.tri-node {{
  position: absolute;
  width: 174px;
  padding: 12px;
  text-align: center;
  border-radius: 14px;
  color: #fffdf8;
  font-weight: 750;
  box-shadow: 0 5px 15px rgba(37,37,45,0.13);
}}
.tri-top {{left: calc(50% - 87px); top: 25px; background: {GOLD}; color: {INK};}}
.tri-left {{left: 9%; bottom: 32px; background: {BLUE};}}
.tri-right {{right: 9%; bottom: 32px; background: {RED};}}
.tri-center {{
  position: absolute;
  left: calc(50% - 80px);
  top: 145px;
  width: 160px;
  text-align: center;
  color: {INK};
  font-size: 13px;
  font-weight: 750;
}}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# 2. 通用函数
# =========================================================

def hero() -> None:
    st.markdown(
        """
<div class="hero">
  <div class="hero-title">数绘梨园</div>
  <div class="hero-sub">京剧剧本多维可视分析平台｜行当识别 · 人物关系 · 主题谱系 · 叙事节奏 · 三元协同</div>
</div>
<div class="gold-divider"></div>
""",
        unsafe_allow_html=True,
    )


def section_title(text: str) -> None:
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


def metric_card(label: str, value: object) -> None:
    st.markdown(
        f"""
<div class="metric-card">
  <div class="metric-label">{label}</div>
  <div class="metric-value">{value}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def fmt_num(value: object, digits: int = 2) -> str:
    try:
        if pd.isna(value):
            return "—"
        return f"{float(value):.{digits}f}"
    except Exception:
        return "—"


def most_common(series: pd.Series, default: str = "未知") -> str:
    s = series.dropna().astype(str)
    s = s[~s.isin(["", "nan", "None"])]
    if s.empty:
        return default
    return str(s.value_counts().index[0])


def normalize_id(series: pd.Series) -> pd.Series:
    return series.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()


def normalize_play_id(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    if "play_id" in out.columns:
        out["play_id"] = normalize_id(out["play_id"])
    return out


def unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    seen: Dict[str, int] = {}
    cols: List[str] = []
    for raw in out.columns:
        col = str(raw)
        count = seen.get(col, 0)
        cols.append(col if count == 0 else f"{col}__{count}")
        seen[col] = count + 1
    out.columns = cols
    return out


def first_existing(df: pd.DataFrame, candidates: Sequence[str]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def first_valid(*values: object, default: str = "未知") -> str:
    for value in values:
        if value is None:
            continue
        try:
            if pd.isna(value):
                continue
        except Exception:
            pass
        text = str(value).strip()
        if text not in ["", "nan", "None", "未知"]:
            return text
    return default


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def find_file(root: Path, candidates: Sequence[str]) -> Optional[Path]:
    root = Path(root)
    for name in candidates:
        direct = root / name
        if direct.exists():
            return direct
        matches = list(root.rglob(name))
        if matches:
            return matches[0]
    return None


def read_table(path: Optional[Path]) -> pd.DataFrame:
    if path is None or not Path(path).exists():
        return pd.DataFrame()
    path = Path(path)
    try:
        if path.suffix.lower() == ".csv":
            try:
                return pd.read_csv(path, encoding="utf-8-sig")
            except UnicodeDecodeError:
                return pd.read_csv(path, encoding="gbk")
        return pd.read_excel(path)
    except Exception as exc:
        st.warning(f"读取失败：{path.name}；原因：{exc}")
        return pd.DataFrame()


@st.cache_data(show_spinner="正在读取第一至第四问结果……")
def load_all_data(root_text: str) -> Tuple[Dict[str, Optional[Path]], Dict[str, pd.DataFrame]]:
    root = Path(root_text)
    paths = {key: find_file(root, names) for key, names in FILE_CANDIDATES.items()}
    data = {
        key: normalize_play_id(unique_columns(read_table(path)))
        for key, path in paths.items()
    }
    return paths, data


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="融合总表", index=False)
    return buffer.getvalue()


# =========================================================
# 3. 表格标准化
# =========================================================

def prepare_role_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    将第一问角色级结果统一为面板字段。

    除性别、年龄、身份外，V5 新增：
    - 性格描述/性格关键词；
    - 唱、白、念、做、打角色级提示或计数；
    - 角色文本/表演提示文本；
    从而能够在面板中直接回答“角色特征—行当对应模式”。
    """
    if df.empty or "play_id" not in df.columns:
        return pd.DataFrame()

    out = df.copy()
    title_col = first_existing(out, ["play_title", "剧目", "剧名"])
    role_col = first_existing(out, ["role", "角色", "角色名"])
    big_col = first_existing(out, ["analysis_hangdang_big", "最终预测行当大类", "hangdang_big", "行当大类"])
    fine_col = first_existing(out, ["analysis_hangdang_fine", "最终建议行当", "hangdang_fine", "细分行当"])
    period_col = first_existing(out, ["analysis_period", "historical_period", "creation_period", "历史时期", "创作年代"])
    gender_col = first_existing(out, ["gender_guess", "性别倾向", "性别"])
    age_col = first_existing(out, ["age_guess", "年龄倾向", "年龄"])
    identity_col = first_existing(out, ["identity_guess", "身份类型", "身份"])
    personality_col = first_existing(out, [
        "personality_guess", "personality", "personality_type",
        "性格倾向", "性格描述", "性格类型", "性格关键词", "trait_keywords",
    ])
    source_col = first_existing(out, ["role_source", "source", "角色来源"])
    conf_col = first_existing(out, ["细分预测置信度", "大类预测置信度", "prediction_confidence", "预测置信度"])
    performance_text_col = first_existing(out, [
        "performance_text", "performance_hint", "role_context", "character_text",
        "角色文本", "角色描述", "表演提示", "唱念做打提示", "feature_text",
    ])

    performance_aliases = {
        "sing_std": ["sing_count", "唱次数", "唱段次数", "唱腔次数", "唱"],
        "speak_std": ["speak_count", "dialogue_count", "白次数", "对白次数", "白"],
        "recite_std": ["recite_count", "念次数", "念白次数", "念"],
        "action_std": ["action_count", "做次数", "动作次数", "舞台动作次数", "做"],
        "martial_std": ["martial_count", "fight_count", "打次数", "武打次数", "开打次数", "打"],
    }

    result = out.copy()
    result["play_id"] = normalize_id(out["play_id"])
    result["play_title_std"] = out[title_col].astype(str) if title_col else result["play_id"]
    result["role_std"] = out[role_col].astype(str) if role_col else "未知角色"
    result["hangdang_big_std"] = out[big_col].fillna("未知").astype(str) if big_col else "未知"
    result["hangdang_fine_std"] = out[fine_col].fillna("未知").astype(str) if fine_col else "未知"
    result["period_std"] = out[period_col].fillna("未知").astype(str) if period_col else "未知"
    result["gender_std"] = out[gender_col].fillna("未知").astype(str) if gender_col else "未知"
    result["age_std"] = out[age_col].fillna("未知").astype(str) if age_col else "未知"
    result["identity_std"] = out[identity_col].fillna("未知身份").astype(str) if identity_col else "未知身份"
    result["personality_std"] = out[personality_col].fillna("未知性格").astype(str) if personality_col else "未知性格"
    result["source_std"] = out[source_col].fillna("未知").astype(str) if source_col else "未知"
    result["confidence_std"] = safe_numeric(out[conf_col]) if conf_col else 0.0
    result["performance_text_std"] = out[performance_text_col].fillna("").astype(str) if performance_text_col else ""

    # 优先读取已计算的角色级次数；若缺失，则从表演提示文本中做保守计数。
    marker_map = {
        "sing_std": r"唱|西皮|二黄|南梆子|高拨子",
        "speak_std": r"白|对白|同白",
        "recite_std": r"念|同念",
        "action_std": r"做|舞|水袖|圆场|上|下|笑|哭",
        "martial_std": r"打|开打|对刀|起霸|趟马|翻身|枪|剑",
    }
    for target, aliases in performance_aliases.items():
        col = first_existing(out, aliases)
        if col:
            result[target] = safe_numeric(out[col])
        elif performance_text_col:
            result[target] = result["performance_text_std"].str.count(marker_map[target]).astype(float)
        else:
            result[target] = 0.0

    result["performance_total_std"] = result[
        ["sing_std", "speak_std", "recite_std", "action_std", "martial_std"]
    ].sum(axis=1)

    raw_big_col = first_existing(out, ["hangdang_big", "原始行当大类"])
    if raw_big_col:
        raw_big = out[raw_big_col].fillna("未知").astype(str)
        result["predicted_flag"] = raw_big.isin(["未标注", "未知", "其他", "", "nan"])
    else:
        result["predicted_flag"] = result["source_std"].str.contains("正文|预测|推断", na=False)

    return result


def prepare_network_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "play_id" not in df.columns:
        return pd.DataFrame()

    aliases = {
        "play_title": ["play_title", "剧目", "剧名"],
        "play_type": ["play_type", "剧目类型", "类型"],
        "node_count": ["node_count", "节点数", "角色数量"],
        "edge_count": ["edge_count", "边数", "关系数量"],
        "density": ["density", "网络密度"],
        "avg_degree": ["avg_degree", "平均度"],
        "avg_clustering": ["avg_clustering", "平均聚类系数", "clustering"],
        "modularity": ["modularity", "模块度"],
        "network_pattern": ["network_pattern", "网络模式", "网络结构类型"],
    }
    result = pd.DataFrame({"play_id": normalize_id(df["play_id"])})
    for target, candidates in aliases.items():
        col = first_existing(df, candidates)
        if col:
            result[target] = df[col]
    return result.drop_duplicates("play_id")


def detect_theme_column(df: pd.DataFrame, theme: str) -> Optional[str]:
    return first_existing(df, [
        f"{theme}_norm", theme, f"{theme}_score", f"{theme}_强度",
        f"主题_{theme}", f"{theme}主题强度",
    ])


def prepare_theme_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "play_id" not in df.columns:
        return pd.DataFrame()

    result = pd.DataFrame({"play_id": normalize_id(df["play_id"])})
    aliases = {
        "play_title": ["play_title", "剧目", "剧名"],
        "play_type": ["play_type", "剧目类型", "类型"],
        "theme_pattern": ["theme_pattern", "主题模式", "主题组合", "主题组合模式"],
        "top1_theme": ["top1_theme", "第一主题", "核心主题"],
        "top2_theme": ["top2_theme", "第二主题", "次级主题"],
        "top3_theme": ["top3_theme", "第三主题", "辅助主题"],
        "theme_cluster": ["theme_cluster", "主题聚类", "cluster"],
    }
    for target, candidates in aliases.items():
        col = first_existing(df, candidates)
        if col:
            result[target] = df[col]

    for theme in THEME_ORDER:
        col = detect_theme_column(df, theme)
        result[f"theme_{theme}"] = safe_numeric(df[col]) if col else 0.0

    if "theme_pattern" not in result.columns:
        top_cols = [c for c in ["top1_theme", "top2_theme", "top3_theme"] if c in result.columns]
        result["theme_pattern"] = result[top_cols].fillna("").astype(str).agg("—".join, axis=1) if top_cols else "未知主题模式"

    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER]
    result["theme_concentration"] = result[theme_cols].max(axis=1)
    result["theme_diversity"] = (result[theme_cols] > 0.25).sum(axis=1)
    return result.drop_duplicates("play_id")


def prepare_narrative_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "play_id" not in df.columns:
        return pd.DataFrame()

    aliases = {
        "play_title": ["play_title", "剧目", "剧名"],
        "play_type": ["play_type", "剧目类型", "类型"],
        "narrative_pattern": ["narrative_pattern", "叙事模式", "典型叙事模式"],
        "scene_count": ["scene_count", "场次数", "场次总数"],
        "intensity_score_mean": ["intensity_score_mean", "平均剧情强度"],
        "intensity_score_max": ["intensity_score_max", "剧情强度峰值"],
        "intensity_score_std": ["intensity_score_std", "剧情强度波动"],
        "rhythm_score_mean": ["rhythm_score_mean", "平均节奏指数"],
        "rhythm_score_max": ["rhythm_score_max", "节奏峰值"],
        "climax_position": ["climax_position", "高潮位置"],
        "sing_ratio_mean": ["sing_ratio_mean", "平均唱占比"],
        "speak_ratio_mean": ["speak_ratio_mean", "平均白占比"],
        "recite_ratio_mean": ["recite_ratio_mean", "平均念占比"],
        "action_ratio_mean": ["action_ratio_mean", "平均做占比"],
        "martial_ratio_mean": ["martial_ratio_mean", "平均打占比"],
        "conflict_density_mean": ["conflict_density_mean", "平均冲突密度"],
        "emotion_score_mean": ["emotion_score_mean", "平均情绪强度"],
        "transition_density_mean": ["transition_density_mean", "平均转折密度"],
    }
    result = pd.DataFrame({"play_id": normalize_id(df["play_id"])})
    for target, candidates in aliases.items():
        col = first_existing(df, candidates)
        if col:
            result[target] = df[col]
    if "narrative_pattern" not in result.columns:
        result["narrative_pattern"] = "未知叙事模式"
    return result.drop_duplicates("play_id")


def prepare_scene_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "play_id" not in df.columns:
        return pd.DataFrame()
    result = df.copy()
    result["play_id"] = normalize_id(result["play_id"])
    return result


def detect_edge_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
    source = first_existing(df, ["role_a", "source", "source_role", "角色A", "角色1", "起点", "from", "from_role"])
    target = first_existing(df, ["role_b", "target", "target_role", "角色B", "角色2", "终点", "to", "to_role"])
    relation = first_existing(df, ["dominant_relation", "relation_type", "relation", "关系类型", "主要关系", "关系"])
    weight = first_existing(df, ["weight", "edge_weight", "interaction_count", "count", "互动次数", "边权重", "边数"])
    return source, target, relation, weight


def prepare_edge_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "play_id" not in df.columns:
        return pd.DataFrame()

    source, target, relation, weight = detect_edge_columns(df)
    result = pd.DataFrame({"play_id": normalize_id(df["play_id"])})
    result["role_a_std"] = df[source].astype(str) if source else ""
    result["role_b_std"] = df[target].astype(str) if target else ""
    result["relation_std"] = df[relation].fillna("普通互动").astype(str) if relation else "普通互动"
    result["weight_std"] = safe_numeric(df[weight]) if weight else 1.0

    title_col = first_existing(df, ["play_title", "剧目", "剧名"])
    type_col = first_existing(df, ["play_type", "剧目类型", "类型"])
    if title_col:
        result["play_title"] = df[title_col]
    if type_col:
        result["play_type"] = df[type_col]
    return result


# =========================================================
# 4. 融合
# =========================================================

def safe_merge(left: pd.DataFrame, right: pd.DataFrame, tag: str) -> pd.DataFrame:
    if right.empty:
        return left
    right = right.copy()
    rename_map = {
        col: f"{col}_{tag}"
        for col in right.columns
        if col != "play_id" and col in left.columns
    }
    right = right.rename(columns=rename_map)
    return left.merge(right, on="play_id", how="outer")


def coalesce_field(df: pd.DataFrame, target: str) -> pd.DataFrame:
    candidates = [c for c in df.columns if c == target or c.startswith(f"{target}_")]
    if not candidates:
        df[target] = ""
        return df
    value = df[candidates[0]]
    for col in candidates[1:]:
        value = value.combine_first(df[col])
    df[target] = value
    return df


def build_relation_features(edge_df: pd.DataFrame) -> pd.DataFrame:
    if edge_df.empty:
        return pd.DataFrame()

    pivot = pd.pivot_table(
        edge_df,
        index="play_id",
        columns="relation_std",
        values="weight_std",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    relation_cols = [c for c in pivot.columns if c != "play_id"]
    if not relation_cols:
        return pd.DataFrame()

    total = pivot[relation_cols].sum(axis=1).replace(0, np.nan)
    result = pivot[["play_id"]].copy()
    for col in relation_cols:
        result[f"rel_{col}"] = pivot[col] / total
    return result.fillna(0)


def build_role_summary(role_df: pd.DataFrame) -> pd.DataFrame:
    if role_df.empty:
        return pd.DataFrame()

    records = []
    for play_id, sub in role_df.groupby("play_id"):
        nonzero_conf = sub["confidence_std"].replace(0, np.nan)
        records.append({
            "play_id": str(play_id),
            "role_count_q1": int(sub["role_std"].nunique()),
            "predicted_role_count": int(sub["predicted_flag"].sum()),
            "avg_prediction_confidence": float(nonzero_conf.mean()) if nonzero_conf.notna().any() else np.nan,
            "top_hangdang": most_common(sub["hangdang_big_std"]),
            "analysis_period": most_common(sub["period_std"]),
            "play_title_q1": most_common(sub["play_title_std"], str(play_id)),
        })
    return pd.DataFrame(records)


def classify_integrated_mode(row: pd.Series) -> str:
    text = " ".join([
        str(row.get("play_type", "")),
        str(row.get("theme_pattern", "")),
        str(row.get("narrative_pattern", "")),
    ])
    if any(word in text for word in ["神仙", "仪式", "祥瑞", "祝颂"]):
        return "仪式聚集—神仙祥瑞—圆满祝颂型"
    if any(word in text for word in ["公案", "权力", "平冤", "审判", "反转"]):
        return "权力中介—公案救助—反转平冤型"
    if any(word in text for word in ["战争", "历史", "忠义", "军情", "出征", "交锋"]):
        return "阵营对抗—忠义家国—军情出征型"
    if any(word in text for word in ["婚恋", "家庭", "礼教", "情感", "相遇", "误会", "悲剧"]):
        return "亲密闭合—婚恋礼教—情感阻隔型"
    return "多线复合—多主题—综合推进型"


def build_fused(data: Dict[str, pd.DataFrame]) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    role = prepare_role_table(data.get("role", pd.DataFrame()))
    network = prepare_network_table(data.get("network", pd.DataFrame()))
    edge = prepare_edge_table(data.get("edge", pd.DataFrame()))
    theme = prepare_theme_table(data.get("theme", pd.DataFrame()))
    narrative = prepare_narrative_table(data.get("narrative", pd.DataFrame()))
    scene = prepare_scene_table(data.get("scene", pd.DataFrame()))

    tables = [
        ("network", network),
        ("theme", theme),
        ("narrative", narrative),
        ("role", build_role_summary(role)),
        ("relation", build_relation_features(edge)),
    ]
    valid = [(tag, table) for tag, table in tables if not table.empty]
    if not valid:
        return pd.DataFrame(), {
            "role": role,
            "network": network,
            "edge": edge,
            "theme": theme,
            "narrative": narrative,
            "scene": scene,
        }

    fused = valid[0][1].copy()
    for tag, table in valid[1:]:
        fused = safe_merge(fused, table, tag)

    fused = coalesce_field(fused, "play_title")
    if "play_title_q1" in fused.columns:
        fused["play_title"] = fused["play_title"].replace("", np.nan).combine_first(fused["play_title_q1"])
    fused["play_title"] = fused["play_title"].replace("", np.nan).fillna(fused["play_id"])

    fused = coalesce_field(fused, "play_type")
    fused["play_type"] = fused["play_type"].replace("", np.nan).fillna("未知类型")

    if "theme_pattern" not in fused.columns:
        candidates = [c for c in fused.columns if c.startswith("theme_pattern")]
        fused["theme_pattern"] = fused[candidates[0]] if candidates else "未知主题模式"
    fused["theme_pattern"] = fused["theme_pattern"].fillna("未知主题模式").astype(str)

    if "narrative_pattern" not in fused.columns:
        candidates = [c for c in fused.columns if c.startswith("narrative_pattern")]
        fused["narrative_pattern"] = fused[candidates[0]] if candidates else "未知叙事模式"
    fused["narrative_pattern"] = fused["narrative_pattern"].fillna("未知叙事模式").astype(str)

    if "analysis_period" not in fused.columns:
        fused["analysis_period"] = "未知"
    fused["analysis_period"] = fused["analysis_period"].fillna("未知").astype(str)

    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    if theme_cols:
        fused["theme_concentration"] = fused[theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).max(axis=1)
    else:
        fused["theme_concentration"] = 0.0

    fused["integrated_mode"] = fused.apply(classify_integrated_mode, axis=1)
    fused["display_name"] = fused["play_title"].astype(str) + "｜" + fused["play_id"].astype(str)

    prepared = {
        "role": role,
        "network": network,
        "edge": edge,
        "theme": theme,
        "narrative": narrative,
        "scene": scene,
    }
    return fused, prepared


def numeric_feature_columns(df: pd.DataFrame) -> List[str]:
    candidates = [
        "node_count", "edge_count", "density", "avg_degree", "avg_clustering", "modularity",
        "theme_concentration", "theme_diversity",
        "intensity_score_mean", "intensity_score_max", "intensity_score_std",
        "rhythm_score_mean", "rhythm_score_max", "climax_position",
        "sing_ratio_mean", "speak_ratio_mean", "recite_ratio_mean",
        "action_ratio_mean", "martial_ratio_mean",
        "conflict_density_mean", "emotion_score_mean", "transition_density_mean",
        *[f"theme_{theme}" for theme in THEME_ORDER],
        *[c for c in df.columns if c.startswith("rel_")],
    ]
    valid = []
    for col in candidates:
        if col in df.columns and safe_numeric(df[col]).std() > 0:
            valid.append(col)
    return list(dict.fromkeys(valid))


def add_pca(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    result = df.copy()
    features = numeric_feature_columns(result)
    if len(result) < 3 or len(features) < 2:
        result["pca_x"] = 0.0
        result["pca_y"] = 0.0
        return result, features

    matrix = result[features].apply(pd.to_numeric, errors="coerce").fillna(0).values
    matrix = StandardScaler().fit_transform(matrix)
    coords = PCA(n_components=2, random_state=42).fit_transform(matrix)
    result["pca_x"] = coords[:, 0]
    result["pca_y"] = coords[:, 1]
    return result, features


# =========================================================
# 5. 图表
# =========================================================

_PLOT_COUNTER = itertools.count()


def show_plot(fig: go.Figure, height: Optional[int] = None, key: Optional[str] = None) -> None:
    """绘制 Plotly 图表，并为每次调用分配唯一 key，避免 DuplicateElementId。"""
    if height:
        fig.update_layout(height=height)
    unique_key = key or f"plotly_chart_{next(_PLOT_COUNTER)}"
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displaylogo": False},
        key=unique_key,
    )


def plot_hangdang_distribution(role_df: pd.DataFrame, title: str = "行当大类分布") -> None:
    if role_df.empty:
        st.info("缺少第一问角色数据。")
        return

    counts = role_df["hangdang_big_std"].value_counts().reset_index()
    counts.columns = ["行当", "数量"]
    counts["排序"] = counts["行当"].apply(lambda x: BIG_ORDER.index(x) if x in BIG_ORDER else 99)
    counts = counts.sort_values("排序")

    fig = go.Figure(go.Bar(
        x=counts["数量"],
        y=counts["行当"],
        orientation="h",
        marker_color=[HANGDANG_COLORS.get(str(x), "#999999") for x in counts["行当"]],
        text=counts["数量"],
        textposition="outside",
    ))
    fig.update_layout(
        title=title,
        margin=dict(l=20, r=20, t=45, b=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="角色数量",
        yaxis_title="",
    )
    show_plot(fig, 330)


def plot_hangdang_status(role_df: pd.DataFrame) -> None:
    if role_df.empty:
        st.info("缺少角色预测结果。")
        return

    work = role_df.copy()
    work["标注状态"] = np.where(work["predicted_flag"], "模型预测", "原始标注")
    grouped = (
        work.groupby(["hangdang_big_std", "标注状态"])
        .size()
        .reset_index(name="角色数量")
    )
    fig = px.bar(
        grouped,
        x="hangdang_big_std",
        y="角色数量",
        color="标注状态",
        barmode="stack",
        color_discrete_map={"原始标注": BLUE, "模型预测": GOLD},
        title="已标注与未标注预测角色分布",
    )
    fig.update_layout(
        xaxis_title="行当大类",
        yaxis_title="角色数量",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    show_plot(fig, 380)


def plot_period_area(role_df: pd.DataFrame) -> None:
    if role_df.empty:
        st.info("缺少时期字段。")
        return

    pivot = pd.crosstab(role_df["period_std"], role_df["hangdang_big_std"])
    order = [p for p in PERIOD_ORDER if p in pivot.index] + [p for p in pivot.index if p not in PERIOD_ORDER]
    pivot = pivot.reindex(order).fillna(0)
    pct = pivot.div(pivot.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    long = pct.reset_index().melt(id_vars="period_std", var_name="行当", value_name="占比")

    fig = px.area(
        long,
        x="period_std",
        y="占比",
        color="行当",
        color_discrete_map=HANGDANG_COLORS,
        title="不同时期角色—行当占比演化",
    )
    fig.update_layout(
        xaxis_title="创作年代或历史时期背景",
        yaxis_title="角色占比",
        yaxis_tickformat=".0%",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    show_plot(fig, 430)


def plot_identity_hangdang(role_df: pd.DataFrame) -> None:
    if role_df.empty:
        return

    grouped = role_df.groupby(["identity_std", "hangdang_big_std"]).size().reset_index(name="角色数量")
    top_identities = grouped.groupby("identity_std")["角色数量"].sum().nlargest(8).index
    grouped = grouped[grouped["identity_std"].isin(top_identities)]

    fig = px.bar(
        grouped,
        x="identity_std",
        y="角色数量",
        color="hangdang_big_std",
        barmode="stack",
        color_discrete_map=HANGDANG_COLORS,
        title="身份类型与行当分类对应关系",
    )
    fig.update_layout(
        xaxis_title="身份类型",
        yaxis_title="角色数量",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    show_plot(fig, 420)


def plot_network(play_id: str, edge_df: pd.DataFrame, role_df: pd.DataFrame, max_edges: int = 70, height: int = 520, title: str = "角色关系网络｜节点颜色表示行当，大小表示中心性") -> None:
    if edge_df.empty:
        st.info("缺少第二问角色关系边表。")
        return

    edges = edge_df[edge_df["play_id"] == str(play_id)].copy()
    if edges.empty:
        st.info("该剧没有可用关系边。")
        return

    edges = edges.sort_values("weight_std", ascending=False).head(max_edges)
    graph = nx.Graph()
    for _, row in edges.iterrows():
        a, b = str(row["role_a_std"]).strip(), str(row["role_b_std"]).strip()
        if not a or not b or a == "nan" or b == "nan":
            continue
        graph.add_edge(a, b, weight=float(row["weight_std"]), relation=str(row["relation_std"]))

    if graph.number_of_nodes() == 0:
        st.info("该剧角色网络为空。")
        return

    role_attr: Dict[str, dict] = {}
    if not role_df.empty:
        roles = role_df[role_df["play_id"] == str(play_id)]
        for _, row in roles.iterrows():
            role_attr[str(row["role_std"])] = {
                "hangdang": str(row["hangdang_big_std"]),
                "identity": str(row["identity_std"]),
            }

    pos = nx.spring_layout(graph, seed=42, weight="weight", k=0.78)
    fig = go.Figure()

    for relation in sorted(set(nx.get_edge_attributes(graph, "relation").values())):
        xs, ys = [], []
        for source, target, data in graph.edges(data=True):
            if data.get("relation") != relation:
                continue
            x0, y0 = pos[source]
            x1, y1 = pos[target]
            xs += [x0, x1, None]
            ys += [y0, y1, None]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines",
            line=dict(width=1.25, color=RELATION_COLORS.get(relation, "#999999")),
            name=relation, hoverinfo="skip",
        ))

    degree = dict(graph.degree(weight="weight"))
    max_degree = max(degree.values()) if degree else 1
    node_x, node_y, node_text, node_hover, node_size, node_color = [], [], [], [], [], []
    for node in graph.nodes():
        x, y = pos[node]
        hangdang = role_attr.get(node, {}).get("hangdang", "未知")
        identity = role_attr.get(node, {}).get("identity", "未知身份")
        node_x.append(x)
        node_y.append(y)
        node_text.append(node)
        node_hover.append(f"角色：{node}<br>行当：{hangdang}<br>身份：{identity}<br>加权度：{degree.get(node, 0):.2f}")
        node_size.append(12 + degree.get(node, 0) / max_degree * 28)
        node_color.append(HANGDANG_COLORS.get(hangdang, "#999999"))

    fig.add_trace(go.Scatter(
        x=node_x, y=node_y,
        mode="markers+text",
        text=node_text if len(node_text) <= 28 else None,
        textposition="top center",
        hovertext=node_hover,
        hoverinfo="text",
        marker=dict(size=node_size, color=node_color, line=dict(width=1.1, color="white")),
        name="角色",
    ))
    fig.update_layout(
        title=title,
        showlegend=True,
        margin=dict(l=5, r=5, t=45, b=5),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,253,248,0.60)",
        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
    )
    show_plot(fig, height)


def plot_relation_composition(play_id: str, edge_df: pd.DataFrame, height: int = 360, title: str = "关系类型构成") -> None:
    if edge_df.empty:
        return
    sub = edge_df[edge_df["play_id"] == str(play_id)]
    if sub.empty:
        return

    grouped = sub.groupby("relation_std")["weight_std"].sum().reset_index()
    fig = px.pie(
        grouped,
        names="relation_std",
        values="weight_std",
        hole=0.48,
        color="relation_std",
        color_discrete_map=RELATION_COLORS,
        title=title,
    )
    show_plot(fig, height)


def plot_network_type_radar(network_df: pd.DataFrame, height: int = 500, title: str = "不同剧目类型网络结构比较") -> None:
    if network_df.empty or "play_type" not in network_df.columns:
        st.info("缺少剧目类型或网络指标。")
        return

    metrics = [c for c in ["node_count", "edge_count", "density", "avg_degree", "avg_clustering", "modularity"] if c in network_df.columns]
    if len(metrics) < 3:
        st.info("可用于网络结构比较的指标不足。")
        return

    grouped = network_df.groupby("play_type")[metrics].mean().fillna(0)
    for col in metrics:
        min_v, max_v = grouped[col].min(), grouped[col].max()
        grouped[col] = (grouped[col] - min_v) / (max_v - min_v) if max_v != min_v else 0.5

    fig = go.Figure()
    for play_type, row in grouped.iterrows():
        values = row.tolist()
        fig.add_trace(go.Scatterpolar(
            r=values + [values[0]],
            theta=[NETWORK_METRIC_LABELS.get(m, m) for m in metrics] + [NETWORK_METRIC_LABELS.get(metrics[0], metrics[0])],
            fill="toself",
            name=str(play_type),
        ))
    fig.update_layout(
        title=title,
        polar=dict(radialaxis=dict(visible=True, range=[0, 1])),
    )
    show_plot(fig, height)


def theme_row_for_play(play_id: str, theme_df: pd.DataFrame) -> Optional[pd.Series]:
    if theme_df.empty:
        return None
    sub = theme_df[theme_df["play_id"] == str(play_id)]
    if sub.empty:
        return None
    return sub.iloc[0]


def plot_theme_flower(play_id: str, theme_df: pd.DataFrame, play_title: str, height: int = 460) -> None:
    row = theme_row_for_play(play_id, theme_df)
    if row is None:
        st.info("该剧缺少第三问主题强度。")
        return

    values = [float(row.get(f"theme_{theme}", 0) or 0) for theme in THEME_ORDER]
    max_value = max(max(values), 0.1)
    colors = [RED, GOLD, BLUE, "#7B3F00", "#C65D7B", "#756AB6", JADE, "#5C2B29", "#D6A84B", "#7B9E87"]

    fig = go.Figure(go.Barpolar(
        r=values,
        theta=THEME_ORDER,
        width=[24] * len(THEME_ORDER),
        marker_color=colors,
        marker_line_color=PAPER,
        marker_line_width=1,
        opacity=0.88,
        hovertemplate="%{theta}<br>强度：%{r:.3f}<extra></extra>",
    ))
    fig.update_layout(
        title=f"《{play_title}》主题花瓣图",
        polar=dict(
            bgcolor="rgba(255,253,248,0.4)",
            radialaxis=dict(visible=True, range=[0, max_value], showticklabels=False),
            angularaxis=dict(direction="clockwise"),
        ),
        showlegend=False,
        margin=dict(l=40, r=40, t=55, b=35),
    )
    show_plot(fig, height)


def plot_theme_type_heatmap(fused: pd.DataFrame) -> None:
    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    if "play_type" not in fused.columns or not theme_cols:
        st.info("缺少剧目类型或主题强度。")
        return

    grouped = fused.groupby("play_type")[theme_cols].mean().fillna(0)
    fig = px.imshow(
        grouped.values,
        x=[c.replace("theme_", "") for c in theme_cols],
        y=grouped.index.astype(str),
        color_continuous_scale="YlOrRd",
        text_auto=".2f",
        title="不同剧目类型主题构成差异",
    )
    show_plot(fig, 470)


def plot_theme_cluster(fused: pd.DataFrame) -> None:
    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    if len(theme_cols) < 2 or len(fused) < 3:
        st.info("主题强度字段不足。")
        return

    matrix = fused[theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).values
    coords = PCA(n_components=2, random_state=42).fit_transform(StandardScaler().fit_transform(matrix))
    plot_df = fused[["play_title", "play_type", "theme_pattern"]].copy()
    plot_df["主题维度1"] = coords[:, 0]
    plot_df["主题维度2"] = coords[:, 1]

    fig = px.scatter(
        plot_df,
        x="主题维度1",
        y="主题维度2",
        color="play_type",
        hover_name="play_title",
        hover_data=["theme_pattern"],
        title="跨剧本主题聚类散点图",
    )
    show_plot(fig, 440)


def plot_theme_cooccurrence(fused: pd.DataFrame) -> None:
    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    if len(theme_cols) < 2:
        st.info("主题强度字段不足。")
        return

    corr = fused[theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).corr().fillna(0)
    fig = px.imshow(
        corr.values,
        x=[c.replace("theme_", "") for c in theme_cols],
        y=[c.replace("theme_", "") for c in theme_cols],
        color_continuous_scale="RdBu_r",
        zmin=-1, zmax=1,
        text_auto=".2f",
        title="主题共现相关矩阵",
    )
    show_plot(fig, 500)


def plot_scene_curve(play_id: str, scene_df: pd.DataFrame) -> None:
    if scene_df.empty:
        st.info("缺少第四问场次叙事指标。")
        return

    sub = scene_df[scene_df["play_id"] == str(play_id)].copy()
    if sub.empty:
        st.info("该剧缺少场次叙事指标。")
        return

    if "scene_order" in sub.columns:
        sub = sub.sort_values("scene_order")

    x_col = "scene_name" if "scene_name" in sub.columns else ("scene_order" if "scene_order" in sub.columns else None)
    x_values = sub[x_col] if x_col else list(range(1, len(sub) + 1))

    fig = go.Figure()
    mapping = [
        ("intensity_score", "剧情强度"),
        ("rhythm_score", "节奏指数"),
        ("conflict_norm_play", "冲突强度"),
        ("emotion_norm_play", "情绪强度"),
        ("transition_norm_play", "转折强度"),
    ]
    for col, label in mapping:
        if col in sub.columns:
            fig.add_trace(go.Scatter(
                x=x_values,
                y=pd.to_numeric(sub[col], errors="coerce").fillna(0),
                mode="lines+markers",
                name=label,
            ))

    fig.update_layout(
        title="锣鼓点式叙事节奏曲线",
        xaxis_title="场次",
        yaxis_title="强度指数",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,253,248,0.55)",
    )
    show_plot(fig, 420)


def plot_performance_stack(play_id: str, scene_df: pd.DataFrame) -> None:
    if scene_df.empty:
        return
    sub = scene_df[scene_df["play_id"] == str(play_id)].copy()
    if sub.empty:
        return

    if "scene_order" in sub.columns:
        sub = sub.sort_values("scene_order")

    x_col = "scene_name" if "scene_name" in sub.columns else ("scene_order" if "scene_order" in sub.columns else None)
    x_values = sub[x_col] if x_col else list(range(1, len(sub) + 1))

    fig = go.Figure()
    for col, label in {
        "sing_count": "唱",
        "speak_count": "白",
        "recite_count": "念",
        "action_count": "做/舞台说明",
        "martial_count": "打/武打动作",
    }.items():
        if col in sub.columns:
            fig.add_trace(go.Scatter(
                x=x_values,
                y=pd.to_numeric(sub[col], errors="coerce").fillna(0),
                mode="lines",
                stackgroup="one",
                name=label,
            ))

    fig.update_layout(
        title="唱白念做打随场次变化",
        xaxis_title="场次",
        yaxis_title="次数",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,253,248,0.55)",
    )
    show_plot(fig, 420)


def plot_relation_theme_heatmap(fused: pd.DataFrame) -> None:
    rel_cols = [c for c in fused.columns if c.startswith("rel_")]
    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    if len(rel_cols) == 0 or len(theme_cols) == 0 or len(fused) < 3:
        st.info("关系类型或主题强度字段不足，无法绘制相关热力图。")
        return

    corr = fused[rel_cols + theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).corr().loc[rel_cols, theme_cols]
    fig = px.imshow(
        corr.values,
        x=[c.replace("theme_", "") for c in theme_cols],
        y=[c.replace("rel_", "") for c in rel_cols],
        color_continuous_scale="RdBu_r",
        zmin=-1, zmax=1,
        text_auto=".2f",
        title="人物关系类型 × 主题强度相关性",
    )
    show_plot(fig, 520)


def plot_sankey(fused: pd.DataFrame) -> None:
    if fused.empty:
        return

    work = fused[["theme_pattern", "play_type", "narrative_pattern"]].fillna("未知").copy()
    top_theme = work["theme_pattern"].value_counts().head(12).index
    top_narr = work["narrative_pattern"].value_counts().head(12).index
    work["theme_pattern"] = np.where(work["theme_pattern"].isin(top_theme), work["theme_pattern"], "其他主题模式")
    work["narrative_pattern"] = np.where(work["narrative_pattern"].isin(top_narr), work["narrative_pattern"], "其他叙事模式")

    link1 = work.groupby(["theme_pattern", "play_type"]).size().reset_index(name="value")
    link2 = work.groupby(["play_type", "narrative_pattern"]).size().reset_index(name="value")

    labels = []
    for value in link1["theme_pattern"].unique():
        labels.append("主题｜" + str(value))
    for value in work["play_type"].unique():
        labels.append("类型｜" + str(value))
    for value in link2["narrative_pattern"].unique():
        labels.append("叙事｜" + str(value))

    labels = list(dict.fromkeys(labels))
    index = {label: i for i, label in enumerate(labels)}
    source, target, value = [], [], []
    for _, row in link1.iterrows():
        source.append(index["主题｜" + str(row["theme_pattern"])])
        target.append(index["类型｜" + str(row["play_type"])])
        value.append(int(row["value"]))
    for _, row in link2.iterrows():
        source.append(index["类型｜" + str(row["play_type"])])
        target.append(index["叙事｜" + str(row["narrative_pattern"])])
        value.append(int(row["value"]))

    fig = go.Figure(data=[go.Sankey(
        node=dict(
            pad=18,
            thickness=18,
            label=labels,
            color="rgba(140,29,24,0.72)",
        ),
        link=dict(
            source=source,
            target=target,
            value=value,
            color="rgba(200,155,60,0.25)",
        ),
    )])
    fig.update_layout(
        title_text="主题组合 → 剧目类型 → 叙事模式",
        font_size=12,
    )
    show_plot(fig, 610)


def plot_mode_radar(fused: pd.DataFrame) -> None:
    metric_map = {
        "网络规模": "edge_count",
        "网络密度": "density",
        "亲属婚恋": "rel_亲属婚恋",
        "君臣权力": "rel_君臣权力",
        "冲突敌对": "rel_冲突敌对",
        "救助协作": "rel_救助协作",
        "主题集中度": "theme_concentration",
        "剧情强度": "intensity_score_mean",
        "节奏指数": "rhythm_score_mean",
        "高潮后置": "climax_position",
    }
    available = {k: v for k, v in metric_map.items() if v in fused.columns}
    if len(available) < 3:
        st.info("可用于雷达图的综合指标不足。")
        return

    grouped = fused.groupby("integrated_mode")[list(available.values())].mean().fillna(0)
    top_modes = fused["integrated_mode"].value_counts().head(5).index
    grouped = grouped.loc[[m for m in top_modes if m in grouped.index]]

    norm = grouped.copy()
    for col in norm.columns:
        min_v, max_v = norm[col].min(), norm[col].max()
        norm[col] = (norm[col] - min_v) / (max_v - min_v) if max_v != min_v else 0.5

    labels = list(available.keys())
    fig = go.Figure()
    for mode, row in norm.iterrows():
        values = row.tolist()
        fig.add_trace(go.Scatterpolar(
            r=values + [values[0]],
            theta=labels + [labels[0]],
            fill="toself",
            name=mode,
            line_color=MODE_COLORS.get(mode, None),
        ))
    fig.update_layout(
        title="典型综合模式雷达图",
        polar=dict(radialaxis=dict(visible=True, range=[0, 1])),
    )
    show_plot(fig, 560)


def plot_overview_scatter(fused: pd.DataFrame) -> None:
    fig = px.scatter(
        fused,
        x="pca_x",
        y="pca_y",
        color="play_type",
        size="edge_count" if "edge_count" in fused.columns else None,
        size_max=24,
        hover_name="play_title",
        hover_data=["theme_pattern", "narrative_pattern", "integrated_mode"],
        title="综合聚类散点图｜关系 × 主题 × 叙事",
    )
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,253,248,0.55)",
        xaxis_title="综合维度1",
        yaxis_title="综合维度2",
    )
    show_plot(fig, 430)


def similar_plays(fused: pd.DataFrame, play_id: str, features: List[str], top_n: int = 8) -> pd.DataFrame:
    if len(features) < 2 or str(play_id) not in fused["play_id"].astype(str).values:
        return pd.DataFrame()

    matrix = fused[features].apply(pd.to_numeric, errors="coerce").fillna(0).values
    if matrix.shape[0] < 2:
        return pd.DataFrame()

    matrix = StandardScaler().fit_transform(matrix)
    sim = cosine_similarity(matrix)
    idx = fused.index[fused["play_id"].astype(str) == str(play_id)][0]

    result = fused[["play_id", "play_title", "play_type", "theme_pattern", "narrative_pattern", "integrated_mode"]].copy()
    result["相似度"] = sim[idx]
    return result[result["play_id"].astype(str) != str(play_id)].sort_values("相似度", ascending=False).head(top_n)


# =========================================================
# 6. 页面渲染
# =========================================================

# =========================================================
# 6. V5 扩展：题目覆盖、查询、对比与自动解读
# =========================================================

st.markdown(
    f"""
<style>
.task-strip {{
  border-left: 5px solid {RED};
  border-radius: 10px;
  padding: 10px 14px;
  margin: 4px 0 12px 0;
  background: linear-gradient(90deg, rgba(140,29,24,0.09), rgba(200,155,60,0.05));
  color: {INK};
  font-weight: 850;
  font-size: 19px;
}}
.query-box {{
  border: 1px solid rgba(200,155,60,0.35);
  background: rgba(255,253,248,0.92);
  border-radius: 14px;
  padding: 12px 14px;
  box-shadow: 0 3px 10px rgba(37,37,45,0.04);
}}
.search-hit {{
  border-left: 4px solid {RED};
  border-radius: 10px;
  padding: 10px 12px;
  margin: 7px 0;
  background: rgba(255,253,248,0.96);
  box-shadow: 0 2px 8px rgba(37,37,45,0.05);
  line-height: 1.65;
}}
.red-mark {{
  background: rgba(185,39,39,0.13);
  color: #b92727;
  border-radius: 3px;
  padding: 0 2px;
  font-weight: 800;
}}
.compact-note {{
  border-radius: 12px;
  background: rgba(49,91,125,0.07);
  border: 1px solid rgba(49,91,125,0.18);
  padding: 10px 12px;
  color: #4d4a46;
  line-height: 1.7;
  font-size: 13px;
}}
.stage-card {{
  border-radius: 16px;
  padding: 16px 18px;
  min-height: 100px;
  background: rgba(255,253,248,0.96);
  border: 1px solid rgba(200,155,60,0.30);
  box-shadow: 0 3px 10px rgba(37,37,45,0.05);
}}
.stage-card-title {{
  font-size: 17px;
  font-weight: 860;
  color: #2b2a31;
  margin-bottom: 8px;
}}
.stage-card-score {{
  font-size: 14px;
  color: #5d554d;
}}
.role-name-large {{
  font-size: 38px;
  font-weight: 900;
  color: {INK};
  letter-spacing: 2px;
  margin-top: 7px;
}}
.role-hangdang {{
  font-size: 25px;
  font-weight: 850;
  margin-top: 5px;
}}
.status-dot {{
  display:inline-block;
  width:9px;
  height:9px;
  border-radius:50%;
  margin-right:6px;
}}

.theme-portrait-card {{
  min-height: 430px;
  border-radius: 16px;
  padding: 28px 30px;
  background:
    radial-gradient(circle at 86% 14%, rgba(200,155,60,0.15), transparent 30%),
    linear-gradient(145deg, rgba(255,253,248,0.99), rgba(247,243,234,0.96));
  border: 1px solid rgba(140,29,24,0.22);
  box-shadow: inset 0 0 0 3px rgba(200,155,60,0.06), 0 5px 16px rgba(37,37,45,0.05);
  display: flex;
  flex-direction: column;
  justify-content: center;
}}
.theme-portrait-title {{
  font-size: 25px;
  line-height: 1.35;
  font-weight: 900;
  color: {RED};
  margin-bottom: 18px;
}}
.theme-portrait-main {{
  font-size: 22px;
  line-height: 2.0;
  color: {INK};
  font-weight: 650;
}}
.theme-portrait-note {{
  margin-top: 20px;
  padding-top: 15px;
  border-top: 1px solid rgba(200,155,60,0.35);
  font-size: 16px;
  line-height: 1.85;
  color: #5e554d;
}}
.compact-figure-title {{
  font-size: 16px;
  font-weight: 850;
  color: {INK};
  margin: 2px 0 4px 0;
}}
</style>
""",
    unsafe_allow_html=True,
)


def task_title(question_no: int, title: str, subtitle: str = "") -> None:
    suffix = f"｜{subtitle}" if subtitle else ""
    st.markdown(
        f'<div class="task-strip">任务一 · 问题{question_no}｜{title}{suffix}</div>',
        unsafe_allow_html=True,
    )


def highlight_keyword(value: object, keyword: str, limit: int = 380) -> str:
    text = str(value)
    if len(text) > limit:
        text = text[:limit] + "…"
    escaped = html.escape(text)
    if not keyword:
        return escaped
    try:
        pattern = re.compile(re.escape(keyword), flags=re.IGNORECASE)
        return pattern.sub(lambda m: f'<span class="red-mark">{html.escape(m.group(0))}</span>', escaped)
    except re.error:
        return escaped


def searchable_columns(df: pd.DataFrame) -> List[str]:
    columns: List[str] = []
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            columns.append(str(col))
    return columns


def render_highlight_search(
    sources: Dict[str, pd.DataFrame],
    key_prefix: str,
    title: str = "标红查询｜输入未知人物、场次、主题或文本关键词",
    max_hits: int = 12,
) -> None:
    with st.expander(f"🔎 {title}", expanded=False):
        available = {name: df for name, df in sources.items() if df is not None and not df.empty}
        if not available:
            st.info("当前模块没有可检索的数据。")
            return

        c1, c2 = st.columns([0.72, 0.28])
        with c1:
            keyword = st.text_input(
                "查询关键词",
                placeholder="例如：程敬思、武旦、忠义、审判、第三场……",
                key=f"{key_prefix}_keyword",
            ).strip()
        with c2:
            source_names = st.multiselect(
                "查询范围",
                list(available.keys()),
                default=list(available.keys()),
                key=f"{key_prefix}_source",
            )

        if not keyword:
            st.caption("输入关键词后，命中内容会以红色突出显示；最多展示前 12 条。")
            return

        hits: List[Tuple[str, int, Dict[str, str]]] = []
        keyword_lower = keyword.lower()
        for source_name in source_names:
            df = available[source_name]
            text_cols = searchable_columns(df)
            if not text_cols:
                continue
            for idx, row in df.iterrows():
                matched: Dict[str, str] = {}
                for col in text_cols:
                    value = row.get(col, "")
                    if keyword_lower in str(value).lower():
                        matched[col] = str(value)
                if matched:
                    hits.append((source_name, int(idx) if isinstance(idx, (int, np.integer)) else 0, matched))
                    if len(hits) >= max_hits:
                        break
            if len(hits) >= max_hits:
                break

        if not hits:
            st.warning("未检索到匹配内容。可尝试角色名简称、剧名、主题词或场次关键词。")
            return

        st.success(f"检索到匹配内容，当前展示 {len(hits)} 条。")
        for source_name, _, matched in hits:
            snippets = []
            for col, value in list(matched.items())[:4]:
                snippets.append(f"<b>{html.escape(str(col))}</b>：{highlight_keyword(value, keyword)}")
            st.markdown(
                f'<div class="search-hit"><b>{html.escape(source_name)}</b><br>{"<br>".join(snippets)}</div>',
                unsafe_allow_html=True,
            )


def role_query_card(role_df: pd.DataFrame, play_id: str, key_prefix: str) -> None:
    sub = role_df[role_df["play_id"] == str(play_id)].copy() if not role_df.empty else pd.DataFrame()
    if sub.empty:
        st.info("当前剧目没有可查询的角色记录。")
        return

    role_names = sorted(sub["role_std"].dropna().astype(str).unique().tolist())
    query = st.text_input(
        "输入该剧人物名字",
        placeholder=f"例如：{role_names[0] if role_names else '角色名'}",
        key=f"{key_prefix}_role_query",
    ).strip()

    if query:
        exact = sub[sub["role_std"].astype(str).str.lower() == query.lower()]
        candidates = exact if not exact.empty else sub[sub["role_std"].astype(str).str.contains(query, case=False, na=False, regex=False)]
    else:
        candidates = sub.sort_values(["predicted_flag", "confidence_std"], ascending=[False, False]).head(1)

    if candidates.empty:
        st.warning("未找到该人物。可输入人物姓名中的一部分。")
        st.caption("本剧可查询人物：" + "、".join(role_names[:20]))
        return

    candidate_names = candidates["role_std"].astype(str).unique().tolist()
    selected_name = candidate_names[0]
    if len(candidate_names) > 1:
        selected_name = st.selectbox("匹配到多个人物，请选择", candidate_names, key=f"{key_prefix}_role_pick")

    role = candidates[candidates["role_std"].astype(str) == selected_name].iloc[0]
    big = str(role.get("hangdang_big_std", "未知"))
    fine = str(role.get("hangdang_fine_std", "未知"))
    color = HANGDANG_COLORS.get(big, RED)
    predicted = bool(role.get("predicted_flag", False))
    confidence = float(role.get("confidence_std", 0) or 0)
    conf_text = fmt_num(confidence, 2) if predicted and confidence > 0 else "原始标注，无需预测置信度"
    source_text = "模型推断" if predicted else str(role.get("source_std", "原始标注"))

    st.markdown(
        f"""
<div class="role-card">
  <div style="font-size:13px;color:#786e63;">人物检索结果｜《{html.escape(str(role.get('play_title_std', '')))}》</div>
  <div class="role-name-large">{html.escape(str(role.get('role_std', '未知角色')))}</div>
  <div class="role-hangdang" style="color:{color};">{html.escape(fine)} · {html.escape(big)}</div>
  <div style="margin-top:14px;">
    <span class="badge">{html.escape(str(role.get('gender_std', '未知')))}</span>
    <span class="badge">{html.escape(str(role.get('age_std', '未知')))}</span>
    <span class="badge">{html.escape(str(role.get('identity_std', '未知身份')))}</span>
  </div>
  <div style="margin-top:15px;font-size:14px;color:#5e554d;line-height:1.8;">
    性格描述：{html.escape(str(role.get('personality_std', '未知性格')))}<br>
    唱/白/念/做/打：{int(role.get('sing_std', 0) or 0)} / {int(role.get('speak_std', 0) or 0)} / {int(role.get('recite_std', 0) or 0)} / {int(role.get('action_std', 0) or 0)} / {int(role.get('martial_std', 0) or 0)}<br>
    历史时期：{html.escape(str(role.get('period_std', '未知')))}<br>
    标注状态：{'未标注角色预测' if predicted else '原始行当标注'}<br>
    预测置信度：{html.escape(conf_text)}<br>
    角色来源：{html.escape(source_text)}
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def build_play_graph(play_id: str, edge_df: pd.DataFrame, max_edges: int = 120) -> nx.Graph:
    graph = nx.Graph()
    if edge_df.empty:
        return graph
    edges = edge_df[edge_df["play_id"] == str(play_id)].copy()
    if edges.empty:
        return graph
    edges = edges.sort_values("weight_std", ascending=False).head(max_edges)
    for _, row in edges.iterrows():
        a, b = str(row["role_a_std"]).strip(), str(row["role_b_std"]).strip()
        if not a or not b or a == "nan" or b == "nan":
            continue
        graph.add_edge(a, b, weight=float(row["weight_std"]), relation=str(row["relation_std"]))
    return graph


def plot_core_role_centrality(play_id: str, edge_df: pd.DataFrame, height: int = 430, title: str = "核心角色中心性 Top 12") -> None:
    graph = build_play_graph(play_id, edge_df)
    if graph.number_of_nodes() == 0:
        st.info("当前剧目无法计算角色中心性。")
        return

    weighted_degree = dict(graph.degree(weight="weight"))
    betweenness = nx.betweenness_centrality(graph, weight="weight", normalized=True)
    rows = []
    for role in graph.nodes():
        rows.append({
            "角色": role,
            "互动强度": weighted_degree.get(role, 0),
            "中介中心性": betweenness.get(role, 0),
        })
    top = pd.DataFrame(rows).sort_values("互动强度", ascending=False).head(12)
    long = top.melt(id_vars="角色", var_name="中心性指标", value_name="数值")
    fig = px.bar(
        long,
        x="数值",
        y="角色",
        color="中心性指标",
        orientation="h",
        barmode="group",
        title=title,
        color_discrete_map={"互动强度": BLUE, "中介中心性": GOLD},
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"})
    show_plot(fig, height)


def plot_network_pair_metrics(network_df: pd.DataFrame, play_a: str, play_b: str, title_a: str, title_b: str, height: int = 410) -> None:
    metrics = [c for c in ["node_count", "edge_count", "density", "avg_degree", "avg_clustering", "modularity"] if c in network_df.columns]
    if len(metrics) < 2:
        st.info("网络指标不足，无法进行两剧对比。")
        return

    rows = network_df[network_df["play_id"].isin([str(play_a), str(play_b)])].copy()
    if len(rows) < 2:
        st.info("至少有一部剧缺少网络指标。")
        return

    all_values = network_df[metrics].apply(pd.to_numeric, errors="coerce").fillna(0)
    mins, maxs = all_values.min(), all_values.max()
    norm = rows[["play_id"] + metrics].copy()
    for metric in metrics:
        values = pd.to_numeric(norm[metric], errors="coerce").fillna(0)
        norm[metric] = (values - mins[metric]) / (maxs[metric] - mins[metric]) if maxs[metric] != mins[metric] else 0.5
    name_map = {str(play_a): title_a, str(play_b): title_b}
    norm["剧目"] = norm["play_id"].astype(str).map(name_map)
    long = norm.melt(id_vars=["play_id", "剧目"], value_vars=metrics, var_name="网络指标", value_name="标准化值")
    long["网络指标"] = long["网络指标"].map(lambda x: NETWORK_METRIC_LABELS.get(x, x))
    fig = px.bar(
        long,
        x="网络指标",
        y="标准化值",
        color="剧目",
        barmode="group",
        title="图5｜两剧网络结构指标对比（全样本标准化）",
        color_discrete_sequence=[BLUE, RED],
    )
    fig.update_yaxes(range=[0, 1.05])
    show_plot(fig, height)


def plot_theme_pair_difference(theme_df: pd.DataFrame, play_a: str, play_b: str, title_a: str, title_b: str) -> None:
    rows = theme_df[theme_df["play_id"].isin([str(play_a), str(play_b)])].copy()
    if len(rows) < 2:
        st.info("至少有一部剧缺少主题强度。")
        return

    records = []
    for play_id, title in [(play_a, title_a), (play_b, title_b)]:
        row = rows[rows["play_id"] == str(play_id)].iloc[0]
        for theme in THEME_ORDER:
            records.append({"剧目": title, "主题": theme, "主题强度": float(row.get(f"theme_{theme}", 0) or 0)})
    fig = px.bar(
        pd.DataFrame(records),
        x="主题",
        y="主题强度",
        color="剧目",
        barmode="group",
        title="两剧主题构成差异",
        color_discrete_sequence=[GOLD, BLUE],
    )
    show_plot(fig, 430)


def normalized_scene_progress(sub: pd.DataFrame) -> pd.Series:
    if len(sub) <= 1:
        return pd.Series([0.0] * len(sub), index=sub.index)
    return pd.Series(np.linspace(0, 1, len(sub)), index=sub.index)


def plot_narrative_pair_curve(scene_df: pd.DataFrame, play_a: str, play_b: str, title_a: str, title_b: str) -> None:
    fig = go.Figure()
    colors = {title_a: BLUE, title_b: RED}
    found = False
    for play_id, title in [(play_a, title_a), (play_b, title_b)]:
        sub = scene_df[scene_df["play_id"] == str(play_id)].copy()
        if sub.empty:
            continue
        if "scene_order" in sub.columns:
            sub = sub.sort_values("scene_order")
        sub["剧情进度"] = normalized_scene_progress(sub)
        for col, label, dash in [
            ("intensity_score", "剧情强度", "solid"),
            ("rhythm_score", "节奏指数", "dot"),
        ]:
            if col in sub.columns:
                found = True
                fig.add_trace(go.Scatter(
                    x=sub["剧情进度"],
                    y=pd.to_numeric(sub[col], errors="coerce").fillna(0),
                    mode="lines+markers",
                    name=f"{title}｜{label}",
                    line=dict(color=colors[title], dash=dash),
                ))
    if not found:
        st.info("两部剧缺少可比较的场次强度数据。")
        return
    fig.update_layout(
        title="两剧叙事节奏对比｜按剧情相对进度对齐",
        xaxis_title="剧情相对进度",
        yaxis_title="指数",
        xaxis_tickformat=".0%",
    )
    show_plot(fig, 440)


def plot_performance_pair(scene_df: pd.DataFrame, play_a: str, play_b: str, title_a: str, title_b: str) -> None:
    mapping = {
        "sing_count": "唱",
        "speak_count": "白",
        "recite_count": "念",
        "action_count": "做",
        "martial_count": "打",
    }
    records = []
    for play_id, title in [(play_a, title_a), (play_b, title_b)]:
        sub = scene_df[scene_df["play_id"] == str(play_id)]
        values = {label: safe_numeric(sub[col]).sum() if col in sub.columns else 0 for col, label in mapping.items()}
        total = sum(values.values()) or 1
        for label, value in values.items():
            records.append({"剧目": title, "表演形式": label, "占比": value / total})
    fig = px.bar(
        pd.DataFrame(records),
        x="表演形式",
        y="占比",
        color="剧目",
        barmode="group",
        title="两剧唱念做打结构对比",
        color_discrete_sequence=[JADE, RED],
    )
    fig.update_yaxes(tickformat=".0%")
    show_plot(fig, 400)


def plot_mode_structure_heatmap(fused: pd.DataFrame) -> None:
    metric_map = {
        "网络密度": "density",
        "平均度": "avg_degree",
        "亲属婚恋": "rel_亲属婚恋",
        "君臣权力": "rel_君臣权力",
        "冲突敌对": "rel_冲突敌对",
        "救助协作": "rel_救助协作",
        "主题集中度": "theme_concentration",
        "剧情强度": "intensity_score_mean",
        "节奏指数": "rhythm_score_mean",
        "高潮后置": "climax_position",
    }
    available = {label: col for label, col in metric_map.items() if col in fused.columns}
    if len(available) < 3:
        st.info("综合指标不足，无法生成协同结构矩阵。")
        return

    grouped = fused.groupby("integrated_mode")[list(available.values())].mean().fillna(0)
    norm = grouped.copy()
    for col in norm.columns:
        min_v, max_v = norm[col].min(), norm[col].max()
        norm[col] = (norm[col] - min_v) / (max_v - min_v) if max_v != min_v else 0.5
    norm.columns = list(available.keys())
    fig = px.imshow(
        norm.values,
        x=norm.columns,
        y=norm.index,
        text_auto=".2f",
        color_continuous_scale="YlOrRd",
        zmin=0,
        zmax=1,
        title="典型协同模式结构特征矩阵",
    )
    show_plot(fig, 470)


def strongest_relation_theme_pair(fused: pd.DataFrame) -> Tuple[str, float]:
    rel_cols = [c for c in fused.columns if c.startswith("rel_")]
    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    if len(rel_cols) == 0 or len(theme_cols) == 0 or len(fused) < 3:
        return "—", 0.0
    corr = fused[rel_cols + theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).corr().loc[rel_cols, theme_cols]
    if corr.empty:
        return "—", 0.0
    stacked = corr.stack().sort_values(ascending=False)
    if stacked.empty:
        return "—", 0.0
    (rel, theme), value = stacked.index[0], float(stacked.iloc[0])
    return f"{rel.replace('rel_', '')} × {theme.replace('theme_', '')}", value


def plot_integrated_mode_share(fused: pd.DataFrame) -> None:
    if fused.empty or "integrated_mode" not in fused.columns:
        st.info("缺少协同模式字段。")
        return
    counts = fused["integrated_mode"].fillna("未知模式").astype(str).value_counts().reset_index()
    counts.columns = ["协同模式", "剧本数量"]
    color_map = {mode: MODE_COLORS.get(mode, "#999999") for mode in counts["协同模式"]}
    fig = px.pie(
        counts,
        names="协同模式",
        values="剧本数量",
        hole=0.52,
        color="协同模式",
        color_discrete_map=color_map,
        title="图1｜典型协同模式占比",
    )
    show_plot(fig, 360)


def render_synergy_insights(fused: pd.DataFrame) -> None:
    rel_cols = [c for c in fused.columns if c.startswith("rel_")]
    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    relation_insight = "关系—主题字段不足，暂不能计算最强关联。"
    if rel_cols and theme_cols and len(fused) >= 3:
        corr = fused[rel_cols + theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).corr().loc[rel_cols, theme_cols]
        stacked = corr.stack().sort_values(ascending=False)
        if not stacked.empty:
            (rel, theme), value = stacked.index[0], stacked.iloc[0]
            relation_insight = f"{rel.replace('rel_', '')} 与 {theme.replace('theme_', '')} 的正向关联最强（r={value:.2f}），说明该类人物关系更常承担相应主题表达。"

    narrative_insight = "叙事模式字段不足，暂不能识别主导主题。"
    if "narrative_pattern" in fused.columns and theme_cols:
        top_narrative = most_common(fused["narrative_pattern"])
        sub = fused[fused["narrative_pattern"].astype(str) == top_narrative]
        if not sub.empty:
            means = sub[theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).mean().sort_values(ascending=False)
            if not means.empty:
                narrative_insight = f"样本中最常见的叙事模式是“{top_narrative}”，其主导主题为“{means.index[0].replace('theme_', '')}”，表明主题选择会影响剧情组织方式。"

    relation_reform = "网络与叙事指标不足，暂不能计算重塑效应。"
    candidates = [c for c in ["density", "avg_degree", "rel_冲突敌对", "rel_亲属婚恋"] if c in fused.columns]
    if "climax_position" in fused.columns and candidates:
        values = {}
        for col in candidates:
            temp = fused[[col, "climax_position"]].apply(pd.to_numeric, errors="coerce").dropna()
            if len(temp) >= 3:
                values[col] = temp.corr().iloc[0, 1]
        if values:
            best = max(values, key=lambda x: abs(values[x]))
            relation_reform = f"高潮位置与“{best.replace('rel_', '')}”的关联相对突出（r={values[best]:.2f}），提示不同叙事节奏会改变人物关系的集中呈现时点。"

    cols = st.columns(3)
    titles = ["关系如何承载主题", "主题如何组织叙事", "叙事如何重塑关系"]
    texts = [relation_insight, narrative_insight, relation_reform]
    for col, title, body in zip(cols, titles, texts):
        with col:
            st.markdown(
                f'<div class="compact-note"><b>{html.escape(title)}</b><br>{html.escape(body)}</div>',
                unsafe_allow_html=True,
            )


def two_play_selectors(fused: pd.DataFrame, key_prefix: str) -> Tuple[str, str, str, str]:
    names = fused["display_name"].tolist()
    c1, c2 = st.columns(2)
    with c1:
        display_a = st.selectbox("对比剧目 A", names, index=0, key=f"{key_prefix}_a")
    with c2:
        display_b = st.selectbox("对比剧目 B", names, index=min(1, len(names) - 1), key=f"{key_prefix}_b")
    id_a, id_b = display_a.split("｜")[-1], display_b.split("｜")[-1]
    title_a = str(fused[fused["play_id"] == id_a].iloc[0]["play_title"])
    title_b = str(fused[fused["play_id"] == id_b].iloc[0]["play_title"])
    return id_a, id_b, title_a, title_b


# =========================================================
# 7. V5 页面渲染：五道题完整覆盖与交互验证
# =========================================================



def question_coverage(items: Sequence[str]) -> None:
    """以无标题标签形式展示本题已覆盖的子任务。"""
    badges = "".join(
        f'<span style="display:inline-block;margin:3px 5px 3px 0;padding:5px 10px;'
        f'border-radius:999px;background:rgba(108,154,139,0.13);'
        f'border:1px solid rgba(108,154,139,0.35);color:{INK};font-size:12px;">✓ {html.escape(item)}</span>'
        for item in items
    )
    st.markdown(
        f'<div class="compact-note">{badges}</div>',
        unsafe_allow_html=True,
    )


def render_completion_matrix() -> None:
    rows = [
        ("问题一", "未标注行当预测；性别/年龄/身份/性格对应；唱念做打对应；细分支；时期演化"),
        ("问题二", "主要角色互动；关系类型；核心角色；不同剧目类型网络结构；两剧差异"),
        ("问题三", "核心主题；主题构成；主题共现；组合模式；跨剧本共性与差异"),
        ("问题四", "关键阶段；剧情起伏；节奏变化；唱念做打；典型叙事模式与跨剧本比较"),
        ("问题五", "关系承载主题；主题组织叙事；叙事重塑关系；协同演化；稳定结构"),
    ]
    with st.expander("查看五道题覆盖矩阵", expanded=False):
        st.dataframe(
            pd.DataFrame(rows, columns=["题目", "面板对应内容"]),
            use_container_width=True,
            hide_index=True,
        )


def _top_categories(series: pd.Series, n: int = 5, invalid: Sequence[str] = ("未知", "未知身份", "未知性格", "nan", "")) -> List[str]:
    clean = series.fillna("").astype(str)
    clean = clean[~clean.isin(invalid)]
    return clean.value_counts().head(n).index.astype(str).tolist()


def plot_role_feature_hangdang_heatmap(role_df: pd.DataFrame) -> None:
    """
    性别、年龄、身份、性格与行当的典型对应模式。
    每个特征行按自身角色数归一化，便于比较。
    """
    if role_df.empty:
        st.info("缺少角色特征数据。")
        return

    feature_specs = [
        ("性别", "gender_std", 2),
        ("年龄", "age_std", 2),
        ("身份", "identity_std", 3),
        ("性格", "personality_std", 3),
    ]
    frames = []
    for label, col, top_n in feature_specs:
        if col not in role_df.columns:
            continue
        categories = _top_categories(role_df[col], top_n)
        if not categories:
            continue
        sub = role_df[role_df[col].astype(str).isin(categories)]
        grouped = sub.groupby([col, "hangdang_big_std"]).size().reset_index(name="数量")
        grouped["特征"] = label + "＝" + grouped[col].astype(str)
        frames.append(grouped[["特征", "hangdang_big_std", "数量"]])

    if not frames:
        st.info("当前第一问结果中缺少可用的性别、年龄、身份或性格字段。")
        return

    long = pd.concat(frames, ignore_index=True)
    pivot = long.pivot_table(index="特征", columns="hangdang_big_std", values="数量", aggfunc="sum", fill_value=0)
    ordered_cols = [c for c in BIG_ORDER if c in pivot.columns] + [c for c in pivot.columns if c not in BIG_ORDER]
    pivot = pivot.reindex(columns=ordered_cols)
    pct = pivot.div(pivot.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)

    fig = px.imshow(
        pct.values,
        x=pct.columns.astype(str),
        y=pct.index.astype(str),
        color_continuous_scale="YlOrRd",
        zmin=0,
        zmax=max(float(pct.values.max()), 0.01),
        text_auto=".0%",
        title="图2｜性别、年龄、身份、性格与行当的典型对应模式",
        labels={"x": "行当大类", "y": "角色特征", "color": "行内占比"},
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=50, b=25),
        coloraxis_colorbar=dict(len=0.72, thickness=13, y=0.5),
    )
    show_plot(fig, 620)


def plot_performance_hangdang_profile(role_df: pd.DataFrame) -> None:
    """展示各行当在唱、白、念、做、打提示上的平均特征。"""
    perf_map = {
        "唱": "sing_std",
        "白": "speak_std",
        "念": "recite_std",
        "做": "action_std",
        "打": "martial_std",
    }
    available = {label: col for label, col in perf_map.items() if col in role_df.columns and safe_numeric(role_df[col]).sum() > 0}
    if role_df.empty or not available:
        st.info("当前第一问结果表未包含角色级唱、白、念、做、打计数或表演提示文本。面板接口已预留，重新生成含这些字段的第一问结果后会自动显示。")
        return

    grouped = role_df.groupby("hangdang_big_std")[list(available.values())].mean().fillna(0)
    grouped = grouped.loc[[x for x in BIG_ORDER if x in grouped.index]]
    norm = grouped.copy()
    for col in norm.columns:
        max_v = norm[col].max()
        norm[col] = norm[col] / max_v if max_v > 0 else 0

    long = norm.reset_index().melt(
        id_vars="hangdang_big_std",
        var_name="表演字段",
        value_name="相对强度",
    )
    reverse = {col: label for label, col in available.items()}
    long["表演形式"] = long["表演字段"].map(reverse)

    fig = px.line(
        long,
        x="表演形式",
        y="相对强度",
        color="hangdang_big_std",
        markers=True,
        color_discrete_map=HANGDANG_COLORS,
        title="图3｜不同角色行当的唱、白、念、做、打特征谱",
        labels={"hangdang_big_std": "行当"},
    )
    fig.update_yaxes(range=[0, 1.05], tickformat=".0%")
    show_plot(fig, 430)


def plot_fine_hangdang_distribution(role_df: pd.DataFrame) -> None:
    """细分行当分布，并区分原始标注与未标注预测。"""
    if role_df.empty or "hangdang_fine_std" not in role_df.columns:
        st.info("缺少细分行当数据。")
        return

    work = role_df.copy()
    work["标注状态"] = np.where(work["predicted_flag"], "模型预测", "原始标注")
    top = work["hangdang_fine_std"].replace(["未知", "nan", ""], np.nan).dropna().value_counts().head(15).index
    work = work[work["hangdang_fine_std"].isin(top)]
    if work.empty:
        st.info("暂无可用细分行当。")
        return

    grouped = work.groupby(["hangdang_fine_std", "标注状态"]).size().reset_index(name="角色数量")
    fig = px.bar(
        grouped,
        x="角色数量",
        y="hangdang_fine_std",
        color="标注状态",
        orientation="h",
        barmode="stack",
        color_discrete_map={"原始标注": BLUE, "模型预测": GOLD},
        title="图4｜细分行当 Top 15：原始标注与模型预测",
        labels={"hangdang_fine_std": "细分行当"},
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"})
    show_plot(fig, 470)


def render_question1_insights(role_df: pd.DataFrame) -> None:
    if role_df.empty:
        return

    def strongest_mapping(col: str, feature_name: str) -> str:
        if col not in role_df.columns:
            return f"{feature_name}字段缺失。"
        sub = role_df[~role_df[col].astype(str).isin(["未知", "未知身份", "未知性格", "", "nan"])]
        if sub.empty:
            return f"{feature_name}有效记录不足。"
        grouped = sub.groupby([col, "hangdang_big_std"]).size().sort_values(ascending=False)
        (feature, hd), count = grouped.index[0], int(grouped.iloc[0])
        return f"{feature_name}中最突出组合为“{feature}—{hd}”（{count}人次）。"

    period_text = "时期演化字段不足。"
    if "period_std" in role_df.columns:
        period_sub = role_df[~role_df["period_std"].astype(str).isin(["未知", "", "nan"])]
        if not period_sub.empty:
            early = period_sub.groupby("period_std")["hangdang_big_std"].agg(lambda x: most_common(x)).to_dict()
            examples = list(early.items())[:3]
            period_text = "；".join([f"{p}以{hd}最常见" for p, hd in examples]) + "。"

    cols = st.columns(3)
    texts = [
        strongest_mapping("identity_std", "身份—行当"),
        strongest_mapping("personality_std", "性格—行当"),
        period_text,
    ]
    titles = ["典型身份对应", "典型性格对应", "时期变化概览"]
    for col, title, body in zip(cols, titles, texts):
        with col:
            st.markdown(
                f'<div class="compact-note"><b>{html.escape(title)}</b><br>{html.escape(body)}</div>',
                unsafe_allow_html=True,
            )


def render_question2_insights(play_id: str, edge_df: pd.DataFrame, network_df: pd.DataFrame) -> None:
    relation_text = "当前剧目缺少关系边。"
    sub = edge_df[edge_df["play_id"] == str(play_id)] if not edge_df.empty else pd.DataFrame()
    if not sub.empty:
        relation = sub.groupby("relation_std")["weight_std"].sum().sort_values(ascending=False)
        if not relation.empty:
            relation_text = f"当前剧目最主要的互动关系是“{relation.index[0]}”，占加权关系总量的 {relation.iloc[0] / max(relation.sum(), 1):.1%}。"

    type_text = "剧目类型网络指标不足。"
    if not network_df.empty and "play_type" in network_df.columns:
        metrics = [c for c in ["edge_count", "density", "avg_clustering"] if c in network_df.columns]
        if metrics:
            grouped = network_df.groupby("play_type")[metrics].mean()
            parts = []
            if "edge_count" in grouped.columns:
                parts.append(f"{grouped['edge_count'].idxmax()}关系规模最大")
            if "density" in grouped.columns:
                parts.append(f"{grouped['density'].idxmax()}网络最紧密")
            if "avg_clustering" in grouped.columns:
                parts.append(f"{grouped['avg_clustering'].idxmax()}局部聚集最强")
            type_text = "；".join(parts) + "。"

    graph = build_play_graph(play_id, edge_df)
    core_text = "核心角色不足。"
    if graph.number_of_nodes() > 0:
        core = max(dict(graph.degree(weight="weight")), key=dict(graph.degree(weight="weight")).get)
        core_text = f"“{core}”的加权互动度最高，是当前剧目关系网络中的主要叙事枢纽。"

    cols = st.columns(3)
    for col, title, body in zip(cols, ["主要关系", "类型差异", "核心角色"], [relation_text, type_text, core_text]):
        with col:
            st.markdown(f'<div class="compact-note"><b>{title}</b><br>{html.escape(body)}</div>', unsafe_allow_html=True)


def plot_theme_pattern_frequency(fused: pd.DataFrame) -> None:
    if fused.empty or "theme_pattern" not in fused.columns:
        st.info("缺少主题组合模式。")
        return
    counts = fused["theme_pattern"].replace(["未知主题模式", "", "nan"], np.nan).dropna().value_counts().head(12).reset_index()
    counts.columns = ["主题组合模式", "剧本数量"]
    if counts.empty:
        st.info("暂无可用主题组合模式。")
        return
    fig = px.bar(
        counts.sort_values("剧本数量"),
        x="剧本数量",
        y="主题组合模式",
        orientation="h",
        color="剧本数量",
        color_continuous_scale="YlOrRd",
        title="图4｜代表性主题组合模式及其出现频次",
    )
    fig.update_coloraxes(showscale=False)
    show_plot(fig, 470)


def render_question3_insights(fused: pd.DataFrame) -> None:
    theme_cols = [f"theme_{theme}" for theme in THEME_ORDER if f"theme_{theme}" in fused.columns]
    if not theme_cols:
        return

    overall = fused[theme_cols].apply(pd.to_numeric, errors="coerce").fillna(0).mean().sort_values(ascending=False)
    common_text = f"跨剧本平均强度最高的主题是“{overall.index[0].replace('theme_', '')}”，构成样本的共同主题底色。"

    diff_text = "剧目类型不足，暂不能判断主题差异。"
    if "play_type" in fused.columns and fused["play_type"].nunique() > 1:
        grouped = fused.groupby("play_type")[theme_cols].mean()
        ranges = (grouped.max() - grouped.min()).sort_values(ascending=False)
        if not ranges.empty:
            diff_text = f"不同剧目类型差异最大的主题是“{ranges.index[0].replace('theme_', '')}”，类型间平均强度跨度为 {ranges.iloc[0]:.2f}。"

    pattern_text = "主题组合模式不足。"
    patterns = fused["theme_pattern"].replace(["未知主题模式", "", "nan"], np.nan).dropna()
    if not patterns.empty:
        pattern_text = f"最具代表性的组合模式为“{patterns.value_counts().index[0]}”，覆盖 {patterns.value_counts().iloc[0]} 部剧本。"

    cols = st.columns(3)
    for col, title, body in zip(cols, ["主题共性", "跨类型差异", "代表性组合"], [common_text, diff_text, pattern_text]):
        with col:
            st.markdown(f'<div class="compact-note"><b>{title}</b><br>{html.escape(body)}</div>', unsafe_allow_html=True)


def infer_narrative_stages(scene_df: pd.DataFrame, play_id: str) -> pd.DataFrame:
    sub = scene_df[scene_df["play_id"] == str(play_id)].copy() if not scene_df.empty else pd.DataFrame()
    if sub.empty:
        return sub
    if "scene_order" in sub.columns:
        sub = sub.sort_values("scene_order")
    sub = sub.reset_index(drop=True)
    sub["剧情进度"] = np.linspace(0, 1, len(sub)) if len(sub) > 1 else 0.0

    existing = first_existing(sub, ["stage", "stage_label", "narrative_stage", "剧情阶段", "叙事阶段"])
    if existing:
        sub["stage_std"] = sub[existing].fillna("发展").astype(str)
    else:
        def base_stage(progress: float) -> str:
            if progress <= 0.15:
                return "开端"
            if progress <= 0.55:
                return "发展"
            if progress <= 0.75:
                return "转折"
            if progress <= 0.90:
                return "高潮"
            return "结局"
        sub["stage_std"] = sub["剧情进度"].map(base_stage)

    intensity_col = first_existing(sub, ["intensity_score", "剧情强度"])
    if intensity_col and len(sub) > 0:
        climax_idx = pd.to_numeric(sub[intensity_col], errors="coerce").fillna(0).idxmax()
        sub.loc[climax_idx, "stage_std"] = "高潮"
    if len(sub) > 0:
        sub.loc[0, "stage_std"] = "开端"
        sub.loc[len(sub) - 1, "stage_std"] = "结局"
    return sub


def render_interactive_stage_tiles(play_id: str, scene_df: pd.DataFrame) -> None:
    """
    将第四问图1改成可多选的“场次方块矩阵”。
    - 每个方块代表一个场次；
    - 颜色代表开端、发展、转折、高潮、结局；
    - 支持单击、框选、套索多选；
    - 下方自动展示选中场次的个例卡片。
    """
    sub = infer_narrative_stages(scene_df, play_id)
    if sub.empty:
        st.info("该剧缺少场次数据，无法生成场次矩阵。")
        return

    scene_col = first_existing(sub, ["scene_name", "scene_marker", "scene_order"])
    preview_col = first_existing(sub, ["scene_text_preview", "scene_text", "场次文本", "文本预览"])
    intensity_col = first_existing(sub, ["intensity_score", "剧情强度"])
    rhythm_col = first_existing(sub, ["rhythm_score", "节奏指数"])

    sub = sub.reset_index(drop=True).copy()
    sub["scene_label"] = (
        sub[scene_col].fillna("").astype(str)
        if scene_col else pd.Series([f"第{i + 1}场" for i in range(len(sub))])
    )
    sub["scene_label"] = [
        label if label and label not in ["nan", "None"] else f"第{i + 1}场"
        for i, label in enumerate(sub["scene_label"].tolist())
    ]
    sub["intensity_std"] = (
        pd.to_numeric(sub[intensity_col], errors="coerce").fillna(0)
        if intensity_col else 0.0
    )
    sub["rhythm_std"] = (
        pd.to_numeric(sub[rhythm_col], errors="coerce").fillna(0)
        if rhythm_col else 0.0
    )
    sub["preview_std"] = (
        sub[preview_col].fillna("").astype(str)
        if preview_col else ""
    )

    # 每行最多放 10 个方块；剧本较长时自动换行，适合页面与截图展示。
    columns_per_row = 10 if len(sub) > 10 else max(len(sub), 1)
    sub["grid_x"] = np.arange(len(sub)) % columns_per_row
    sub["grid_y"] = -(np.arange(len(sub)) // columns_per_row)

    color_map = {
        "开端": "#315B7D",
        "发展": "#6C9A8B",
        "转折": "#C89B3C",
        "高潮": "#8C1D18",
        "结局": "#8B6F9C",
        "单场": "#999999",
    }
    marker_colors = [color_map.get(str(stage), "#999999") for stage in sub["stage_std"]]

    customdata = np.column_stack([
        np.arange(len(sub)),
        sub["stage_std"].astype(str),
        sub["scene_label"].astype(str),
        sub["intensity_std"].astype(float),
        sub["rhythm_std"].astype(float),
        sub["preview_std"].astype(str).str.slice(0, 180),
    ])

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=sub["grid_x"],
        y=sub["grid_y"],
        mode="markers+text",
        text=sub["scene_label"],
        textposition="middle center",
        customdata=customdata,
        marker=dict(
            symbol="square",
            size=54,
            color=marker_colors,
            line=dict(color="rgba(255,255,255,0.95)", width=2),
        ),
        selected=dict(marker=dict(opacity=1.0, line=dict(color="#C89B3C", width=5))),
        unselected=dict(marker=dict(opacity=0.48)),
        hovertemplate=(
            "<b>%{customdata[2]}</b><br>"
            "叙事阶段：%{customdata[1]}<br>"
            "剧情强度：%{customdata[3]:.2f}<br>"
            "节奏指数：%{customdata[4]:.2f}<br>"
            "%{customdata[5]}<extra></extra>"
        ),
        name="场次",
        showlegend=False,
    ))

    # 单独添加图例，不影响点选索引。
    for stage, color in color_map.items():
        if stage not in sub["stage_std"].astype(str).unique():
            continue
        fig.add_trace(go.Scatter(
            x=[None], y=[None], mode="markers",
            marker=dict(symbol="square", size=13, color=color),
            name=stage,
            hoverinfo="skip",
        ))

    row_count = int(np.ceil(len(sub) / columns_per_row))
    chart_height = max(260, 125 + row_count * 90)
    fig.update_layout(
        title="图1｜叙事场次矩阵：点击或框选多个场次进行个例分析",
        height=chart_height,
        margin=dict(l=20, r=20, t=65, b=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,253,248,0.58)",
        legend=dict(
            orientation="h",
            yanchor="bottom", y=1.01,
            xanchor="right", x=1,
            title="叙事阶段",
        ),
        dragmode="select",
        clickmode="event+select",
    )
    fig.update_xaxes(
        visible=False,
        range=[-0.65, max(columns_per_row - 0.35, 0.65)],
        fixedrange=True,
    )
    fig.update_yaxes(
        visible=False,
        range=[-row_count + 0.35, 0.65],
        fixedrange=True,
    )

    st.caption("操作：直接点击单个方块；使用框选或套索可一次选择多个场次。再次框选可更新个例集合。")
    try:
        event = st.plotly_chart(
            fig,
            width="stretch",
            key=f"q4_stage_tiles_{play_id}",
            on_select="rerun",
            selection_mode=("points", "box", "lasso"),
            config={"displaylogo": False, "scrollZoom": False},
        )
        selected_indices = list(event.selection.point_indices) if event and event.selection else []
    except TypeError:
        # 兼容较旧 Streamlit：图仍可显示，但不支持图中点选。
        st.plotly_chart(
            fig,
            use_container_width=True,
            key=f"q4_stage_tiles_fallback_{play_id}",
            config={"displaylogo": False, "scrollZoom": False},
        )
        selected_indices = []
        st.warning("当前 Streamlit 版本较旧，图中多选未启用。请升级：python -m pip install -U streamlit")

    # 未选择时，默认展示每个阶段剧情强度最高的代表场次。
    if selected_indices:
        selected = sub.iloc[[i for i in selected_indices if 0 <= i < len(sub)]].copy()
        section_title(f"已选择 {len(selected)} 个场次｜个例展示")
    else:
        representative_rows = []
        for stage in ["开端", "发展", "转折", "高潮", "结局"]:
            part = sub[sub["stage_std"].astype(str) == stage]
            if part.empty:
                continue
            representative_rows.append(part.sort_values("intensity_std", ascending=False).iloc[0])
        selected = pd.DataFrame(representative_rows).drop_duplicates(subset=["scene_label"])
        section_title("代表性场次｜点击上方方块可改为自定义多选")

    if selected.empty:
        return

    # 最多展示 10 个，避免页面过长；每行 5 张小卡片。
    selected = selected.head(10).reset_index(drop=True)
    for start in range(0, len(selected), 5):
        batch = selected.iloc[start:start + 5]
        cols = st.columns(len(batch), gap="small")
        for col, (_, row) in zip(cols, batch.iterrows()):
            stage = str(row.get("stage_std", "发展"))
            scene_name = str(row.get("scene_label", "关键场次"))
            intensity_value = fmt_num(row.get("intensity_std", np.nan), 2)
            rhythm_value = fmt_num(row.get("rhythm_std", np.nan), 2)
            stage_color = color_map.get(stage, "#999999")
            with col:
                card_html = (
                    f'<div class="stage-card" style="border-top:4px solid {stage_color};">'
                    f'<div class="stage-card-title">{html.escape(stage)}｜{html.escape(scene_name)}</div>'
                    f'<div class="stage-card-score">剧情强度：{intensity_value}</div>'
                    f'<div class="stage-card-score">节奏指数：{rhythm_value}</div>'
                    f'</div>'
                )
                st.markdown(card_html, unsafe_allow_html=True)

    with st.expander("查看所选场次的文本摘要", expanded=False):
        show_cols = ["scene_label", "stage_std", "intensity_std", "rhythm_std", "preview_std"]
        display_df = selected[show_cols].rename(columns={
            "scene_label": "场次",
            "stage_std": "叙事阶段",
            "intensity_std": "剧情强度",
            "rhythm_std": "节奏指数",
            "preview_std": "文本摘要",
        })
        st.dataframe(display_df, use_container_width=True, hide_index=True)


def plot_narrative_stage_timeline(play_id: str, scene_df: pd.DataFrame) -> None:
    sub = infer_narrative_stages(scene_df, play_id)
    if sub.empty:
        st.info("该剧缺少场次数据，无法识别关键阶段。")
        return

    scene_col = first_existing(sub, ["scene_name", "scene_marker", "scene_order"])
    scene_names = sub[scene_col].astype(str) if scene_col else pd.Series([f"场次{i+1}" for i in range(len(sub))])
    intensity_col = first_existing(sub, ["intensity_score", "剧情强度"])
    intensity = pd.to_numeric(sub[intensity_col], errors="coerce").fillna(0) if intensity_col else pd.Series([1.0] * len(sub))
    sizes = 14 + 24 * (intensity - intensity.min()) / (intensity.max() - intensity.min()) if intensity.max() != intensity.min() else pd.Series([20] * len(sub))
    preview_col = first_existing(sub, ["scene_text_preview", "scene_text", "场次文本", "文本预览"])
    hover = sub[preview_col].fillna("").astype(str).str.slice(0, 140) if preview_col else ""

    color_map = {"开端": "#315B7D", "发展": "#6C9A8B", "转折": "#C89B3C", "高潮": "#8C1D18", "结局": "#8B6F9C", "单场": "#999999"}
    plot_df = pd.DataFrame({
        "剧情进度": sub["剧情进度"].values,
        "轴线": [1] * len(sub),
        "叙事阶段": sub["stage_std"].astype(str).values,
        "节点大小": np.asarray(sizes),
        "场次": scene_names.values,
        "场次标签": [str(x) for x in scene_names.values],
        "文本摘要": list(hover) if not isinstance(hover, str) else [hover] * len(sub),
    })
    fig = px.scatter(
        plot_df,
        x="剧情进度",
        y="轴线",
        color="叙事阶段",
        size="节点大小",
        text="场次标签",
        hover_name="场次",
        hover_data={"剧情进度": ":.0%", "文本摘要": True, "轴线": False, "节点大小": False, "场次标签": False},
        color_discrete_map=color_map,
        title="图1｜叙事关键阶段时间轴",
    )
    fig.update_traces(marker=dict(line=dict(width=1.5, color="white")), textposition="bottom center")
    fig.update_yaxes(visible=False)
    fig.update_xaxes(tickformat=".0%", range=[-0.03, 1.03], title="剧情进度")
    show_plot(fig, 250)


def render_key_stage_cards(play_id: str, scene_df: pd.DataFrame) -> None:
    sub = infer_narrative_stages(scene_df, play_id)
    if sub.empty:
        return
    intensity_col = first_existing(sub, ["intensity_score", "剧情强度"])
    stage_order = ["开端", "发展", "转折", "高潮", "结局"]
    rows = []
    for stage in stage_order:
        part = sub[sub["stage_std"].astype(str) == stage].copy()
        if part.empty:
            continue
        if intensity_col:
            part["__score"] = pd.to_numeric(part[intensity_col], errors="coerce").fillna(0)
            row = part.sort_values("__score", ascending=False).iloc[0]
        else:
            row = part.iloc[0]
        rows.append(row)
    if not rows:
        return
    chosen = pd.DataFrame(rows)
    scene_col = first_existing(chosen, ["scene_name", "scene_marker", "scene_order"])
    cols = st.columns(len(chosen), gap="small")
    for col, (_, row) in zip(cols, chosen.iterrows()):
        scene_name = str(row.get(scene_col, "关键场次")) if scene_col else "关键场次"
        stage = str(row.get("stage_std", "发展"))
        score = fmt_num(row.get(intensity_col, np.nan), 2) if intensity_col else "—"
        with col:
            st.markdown(
                f'<div class="stage-card">'
                f'<div class="stage-card-title">{html.escape(stage)}｜{html.escape(scene_name)}</div>'
                f'<div class="stage-card-score">剧情强度：{score}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )


def plot_narrative_pattern_frequency(fused: pd.DataFrame) -> None:
    if fused.empty or "narrative_pattern" not in fused.columns:
        st.info("缺少叙事模式字段。")
        return
    counts = fused["narrative_pattern"].replace(["未知叙事模式", "", "nan"], np.nan).dropna().value_counts().head(12).reset_index()
    counts.columns = ["叙事模式", "剧本数量"]
    if counts.empty:
        st.info("暂无可用叙事模式。")
        return
    fig = px.bar(
        counts.sort_values("剧本数量"),
        x="剧本数量",
        y="叙事模式",
        orientation="h",
        color="剧本数量",
        color_continuous_scale="Bluered",
        title="图4｜典型叙事模式及其出现频次",
    )
    fig.update_coloraxes(showscale=False)
    show_plot(fig, 460)


def plot_narrative_type_structure(fused: pd.DataFrame) -> None:
    metric_map = {
        "剧情强度": "intensity_score_mean",
        "节奏指数": "rhythm_score_mean",
        "高潮位置": "climax_position",
        "唱占比": "sing_ratio_mean",
        "白占比": "speak_ratio_mean",
        "念占比": "recite_ratio_mean",
        "做占比": "action_ratio_mean",
        "打占比": "martial_ratio_mean",
    }
    available = {label: col for label, col in metric_map.items() if col in fused.columns}
    if "play_type" not in fused.columns or len(available) < 3:
        st.info("剧目类型或叙事结构指标不足。")
        return

    grouped = fused.groupby("play_type")[list(available.values())].mean().fillna(0)
    norm = grouped.copy()
    for col in norm.columns:
        min_v, max_v = norm[col].min(), norm[col].max()
        norm[col] = (norm[col] - min_v) / (max_v - min_v) if max_v != min_v else 0.5
    norm.columns = list(available.keys())

    fig = px.imshow(
        norm.values,
        x=norm.columns,
        y=norm.index.astype(str),
        color_continuous_scale="YlGnBu",
        zmin=0,
        zmax=1,
        text_auto=".2f",
        title="图5｜不同剧目类型的叙事结构特征矩阵",
    )
    show_plot(fig, 470)


def render_question4_insights(fused: pd.DataFrame, play_id: str) -> None:
    current = fused[fused["play_id"] == str(play_id)]
    current_text = "当前剧目缺少剧本级叙事指标。"
    if not current.empty:
        row = current.iloc[0]
        current_text = (
            f"当前剧目归入“{row.get('narrative_pattern', '未知叙事模式')}”，"
            f"平均剧情强度 {fmt_num(row.get('intensity_score_mean', np.nan))}，"
            f"高潮相对位置 {fmt_num(row.get('climax_position', np.nan))}。"
        )

    pattern_text = "典型叙事模式不足。"
    patterns = fused["narrative_pattern"].replace(["未知叙事模式", "", "nan"], np.nan).dropna()
    if not patterns.empty:
        pattern_text = f"样本中最常见的叙事模式为“{patterns.value_counts().index[0]}”，共 {patterns.value_counts().iloc[0]} 部剧本。"

    type_text = "不同剧目类型指标不足。"
    if "play_type" in fused.columns and "climax_position" in fused.columns:
        temp = fused.groupby("play_type")["climax_position"].mean().dropna()
        if not temp.empty:
            type_text = f"平均高潮最靠后的剧目类型是“{temp.idxmax()}”（{temp.max():.2f}），说明其铺垫与延宕更长。"

    cols = st.columns(3)
    for col, title, body in zip(cols, ["单剧结构", "典型模式", "类型差异"], [current_text, pattern_text, type_text]):
        with col:
            st.markdown(f'<div class="compact-note"><b>{title}</b><br>{html.escape(body)}</div>', unsafe_allow_html=True)


def plot_narrative_relation_heatmap(fused: pd.DataFrame) -> None:
    rel_cols = [c for c in fused.columns if c.startswith("rel_")]
    if "narrative_pattern" not in fused.columns or not rel_cols:
        st.info("缺少叙事模式或关系类型指标。")
        return

    top_patterns = fused["narrative_pattern"].value_counts().head(10).index
    sub = fused[fused["narrative_pattern"].isin(top_patterns)]
    grouped = sub.groupby("narrative_pattern")[rel_cols].mean().fillna(0)
    fig = px.imshow(
        grouped.values,
        x=[c.replace("rel_", "") for c in rel_cols],
        y=grouped.index.astype(str),
        color_continuous_scale="PuBuGn",
        text_auto=".2f",
        title="图3｜不同叙事方式下的人物关系呈现差异",
        labels={"x": "人物关系类型", "y": "叙事模式", "color": "平均占比"},
    )
    show_plot(fig, 520)


def plot_period_mode_evolution(fused: pd.DataFrame) -> None:
    if "analysis_period" not in fused.columns or "integrated_mode" not in fused.columns:
        st.info("缺少时期或综合模式字段。")
        return

    sub = fused[~fused["analysis_period"].astype(str).isin(["未知", "", "nan"])].copy()
    if sub.empty or sub["analysis_period"].nunique() < 2:
        st.info("有效时期类别不足，无法展示协同演化。")
        return

    counts = sub.groupby(["analysis_period", "integrated_mode"]).size().reset_index(name="剧本数量")
    totals = counts.groupby("analysis_period")["剧本数量"].transform("sum")
    counts["占比"] = counts["剧本数量"] / totals
    existing_periods = counts["analysis_period"].astype(str).unique().tolist()
    order = [p for p in PERIOD_ORDER if p in existing_periods] + [p for p in existing_periods if p not in PERIOD_ORDER]
    counts["analysis_period"] = pd.Categorical(counts["analysis_period"], categories=order, ordered=True)
    counts = counts.sort_values("analysis_period")

    fig = px.area(
        counts,
        x="analysis_period",
        y="占比",
        color="integrated_mode",
        color_discrete_map=MODE_COLORS,
        title="图4｜不同时期关系—主题—叙事协同模式的演化",
        labels={"analysis_period": "历史时期", "integrated_mode": "协同模式"},
    )
    fig.update_yaxes(tickformat=".0%")
    show_plot(fig, 470)


def render_question5_conclusion(fused: pd.DataFrame) -> None:
    top_mode = most_common(fused["integrated_mode"])
    top_share = (fused["integrated_mode"].astype(str) == top_mode).mean() if len(fused) else 0
    text = (
        f"当前筛选样本中，“{top_mode}”占比最高（{top_share:.1%}）。"
        "面板将其拆分为关系承载主题、主题组织叙事、叙事重塑关系和时期协同演化四条证据链，"
        "避免与第四问的单剧叙事曲线重复。"
    )
    st.markdown(
        f'<div class="compact-note"><b>第五问综合结论</b><br>{html.escape(text)}</div>',
        unsafe_allow_html=True,
    )



def render_role_card(role_df: pd.DataFrame, play_id: str) -> None:
    role_query_card(role_df, play_id, "overview_role")


def render_mode_cards() -> None:
    cols = st.columns(4)
    modes = [
        "阵营对抗—忠义家国—军情出征型",
        "权力中介—公案救助—反转平冤型",
        "亲密闭合—婚恋礼教—情感阻隔型",
        "仪式聚集—神仙祥瑞—圆满祝颂型",
    ]
    for col, mode in zip(cols, modes):
        with col:
            st.markdown(
                f"""
<div class="mode-card" style="background:{MODE_COLORS[mode]};">
  <h4>{mode}</h4>
  <p>{MODE_EXPLAIN[mode]}</p>
</div>
""",
                unsafe_allow_html=True,
            )


def render_triangle() -> None:
    st.markdown(
        """
<div class="triangle">
  <div class="tri-node tri-top">主题表达<br><span style="font-size:12px;">主题组合与意义结构</span></div>
  <div class="tri-node tri-left">人物关系<br><span style="font-size:12px;">角色网络与关系类型</span></div>
  <div class="tri-node tri-right">叙事结构<br><span style="font-size:12px;">阶段节奏与剧情推进</span></div>
  <div class="tri-center">承载主题<br>组织叙事<br>重塑关系</div>
</div>
""",
        unsafe_allow_html=True,
    )


def render_system_overview(
    fused: pd.DataFrame,
    prepared: Dict[str, pd.DataFrame],
    selected_play_id: str,
    selected_play_title: str,
) -> None:
    section_title("系统总览｜任务一五问联动驾驶舱")
    render_completion_matrix()
    role_df, edge_df = prepared["role"], prepared["edge"]
    theme_df, scene_df = prepared["theme"], prepared["scene"]

    cols = st.columns(6)
    values = [
        ("剧本数量", len(fused)),
        ("角色数量", len(role_df) if not role_df.empty else "—"),
        ("预测角色", int(role_df["predicted_flag"].sum()) if not role_df.empty else "—"),
        ("关系边数", len(edge_df) if not edge_df.empty else "—"),
        ("主题模式", fused["theme_pattern"].nunique()),
        ("叙事模式", fused["narrative_pattern"].nunique()),
    ]
    for col, (label, value) in zip(cols, values):
        with col:
            metric_card(label, value)

    left, right = st.columns(2)
    with left:
        task_title(1, "角色行当推断", "已标注与预测角色概览")
        plot_hangdang_status(role_df)
    with right:
        task_title(2, "角色关系网络", "当前剧目互动结构")
        plot_network(selected_play_id, edge_df, role_df, max_edges=42)

    left, right = st.columns(2)
    with left:
        task_title(3, "核心主题结构", "当前剧目主题花瓣")
        plot_theme_flower(selected_play_id, theme_df, selected_play_title)
    with right:
        task_title(4, "叙事结构分析", "当前剧目节奏曲线")
        plot_scene_curve(selected_play_id, scene_df)

    task_title(5, "三元关联机制", "人物关系—主题表达—叙事结构")
    a, b = st.columns([1, 1.3])
    with a:
        plot_relation_theme_heatmap(fused)
    with b:
        plot_sankey(fused)


def render_question1(fused: pd.DataFrame, prepared: Dict[str, pd.DataFrame], play_id: str) -> None:
    task_title(1, "角色行当推断", "未标注预测、特征对应、细分行当与时代演化")
    question_coverage([
        "推断未标注角色的大类与细分行当",
        "分析性别、年龄、身份、性格与行当对应",
        "分析唱、白、念、做、打表演提示",
        "呈现预测结果与置信度",
        "比较不同时期角色—行当变化",
    ])
    role_df = prepared["role"]

    # 图1 + 人物查询
    c1, c2 = st.columns([1.18, 0.82], gap="small")
    with c1:
        plot_hangdang_status(role_df)
    with c2:
        role_query_card(role_df, play_id, "q1")

    # 图2、图3
    c3, c4 = st.columns([1.05, 0.95], gap="small")
    with c3:
        left_pad, mid_plot, right_pad = st.columns([0.08, 0.84, 0.08])
        with mid_plot:
            plot_role_feature_hangdang_heatmap(role_df)
    with c4:
        plot_performance_hangdang_profile(role_df)

    # 图4、图5
    c5, c6 = st.columns([0.88, 1.12], gap="small")
    with c5:
        plot_fine_hangdang_distribution(role_df)
    with c6:
        plot_period_area(role_df)

    render_question1_insights(role_df)

    render_highlight_search(
        {"人物与行当结果": role_df},
        "q1_search",
        "问题一标红查询｜检索人物、性别、年龄、身份、性格、表演提示、行当与时期",
    )


def render_question2(fused: pd.DataFrame, prepared: Dict[str, pd.DataFrame], play_id: str) -> None:
    task_title(2, "角色关系网络", "人物互动与网络结构")
    question_coverage([
        "识别主要角色之间的互动关系",
        "构建加权角色关系网络",
        "识别核心角色与叙事枢纽",
        "比较历史戏、家庭戏、公案戏等网络结构",
        "支持两部剧的直接差异比较",
    ])
    edge_df, role_df, network_df = prepared["edge"], prepared["role"], prepared["network"]

    main_col, side_col = st.columns([1, 1], gap="small")
    with main_col:
        plot_network(
            play_id, edge_df, role_df, max_edges=85, height=500,
            title="图1｜当前剧目主要角色互动网络",
        )
    with side_col:
        plot_relation_composition(
            play_id, edge_df, height=310, title="图2｜关系类型构成",
        )
        plot_core_role_centrality(
            play_id, edge_df, height=310, title="图3｜核心角色中心性 Top 12",
        )

    render_question2_insights(play_id, edge_df, network_df)

    compare_left, compare_right = st.columns(2, gap="small")
    with compare_left:
        plot_network_type_radar(
            network_df, height=390, title="图4｜不同剧目类型网络结构",
        )
    with compare_right:
        id_a, id_b, title_a, title_b = two_play_selectors(fused, "q2_compare")
        plot_network_pair_metrics(
            network_df, id_a, id_b, title_a, title_b, height=390,
        )

    with st.expander("展开查看两部对比剧目的关系网络", expanded=False):
        a, b = st.columns(2, gap="small")
        with a:
            plot_network(
                id_a, edge_df, role_df, max_edges=45, height=345,
                title=f"《{title_a}》关系网络",
            )
        with b:
            plot_network(
                id_b, edge_df, role_df, max_edges=45, height=345,
                title=f"《{title_b}》关系网络",
            )

    render_highlight_search(
        {"角色关系边": edge_df, "网络指标": network_df},
        "q2_search",
        "问题二标红查询｜检索人物关系、剧目类型、核心角色与网络模式",
    )


def render_question3(fused: pd.DataFrame, prepared: Dict[str, pd.DataFrame], play_id: str, play_title: str) -> None:
    task_title(3, "核心主题结构", "主题构成与组合模式")
    question_coverage([
        "提取单剧核心主题",
        "展示多维主题构成",
        "分析主题共现与组合方式",
        "识别代表性主题组合模式",
        "比较不同剧本与不同类型的共性和差异",
    ])
    theme_df = prepared["theme"]

    a, b = st.columns([1, 1], gap="medium")
    with a:
        plot_theme_flower(play_id, theme_df, play_title, height=430)
    with b:
        row = theme_row_for_play(play_id, theme_df)
        if row is not None:
            top1 = first_valid(row.get("top1_theme", None), default="未知")
            top2 = first_valid(row.get("top2_theme", None), default="未知")
            top3 = first_valid(row.get("top3_theme", None), default="未知")
            pattern = first_valid(row.get("theme_pattern", None), default="未知主题模式")
            concentration = fmt_num(row.get("theme_concentration", np.nan), 2)
            interpretation = (
                f"该剧以“{top1}”为核心，以“{top2}”和“{top3}”作为辅助主题，"
                f"共同形成“{pattern}”。主题集中度为 {concentration}，"
                "可判断其为单核心主导还是多主题并行。"
            )
            st.markdown(
                f"""
<div class="theme-portrait-card">
  <div class="theme-portrait-title">《{html.escape(play_title)}》主题画像</div>
  <div class="theme-portrait-main">
    核心主题：<b>{html.escape(top1)}</b><br>
    次级主题：<b>{html.escape(top2)}</b><br>
    辅助主题：<b>{html.escape(top3)}</b><br>
    主题模式：<b>{html.escape(pattern)}</b><br>
    主题集中度：<b>{concentration}</b>
  </div>
  <div class="theme-portrait-note">{html.escape(interpretation)}</div>
</div>
""",
                unsafe_allow_html=True,
            )
        else:
            st.info("该剧暂无主题画像数据。")

    render_question3_insights(fused)

    # 图2、图3：主题共现与类型差异
    c1, c2 = st.columns(2, gap="small")
    with c1:
        plot_theme_cooccurrence(fused)
    with c2:
        plot_theme_type_heatmap(fused)

    # 图4、图5：代表性组合与两剧比较
    c3, c4 = st.columns([0.9, 1.1], gap="small")
    with c3:
        plot_theme_pattern_frequency(fused)
    with c4:
        id_a, id_b, title_a, title_b = two_play_selectors(fused, "q3_compare")
        plot_theme_pair_difference(theme_df, id_a, id_b, title_a, title_b)

    with st.expander("补充查看跨剧本主题聚类", expanded=False):
        plot_theme_cluster(fused)

    render_highlight_search(
        {"主题结果": theme_df, "融合剧目画像": fused},
        "q3_search",
        "问题三标红查询｜检索核心主题、主题组合、共性差异与代表剧目",
    )


def render_question4(fused: pd.DataFrame, prepared: Dict[str, pd.DataFrame], play_id: str) -> None:
    task_title(4, "叙事结构分析", "关键阶段与节奏模式")
    question_coverage([
        "依据表演标记和剧本内容识别关键阶段",
        "刻画剧情起伏与节奏变化",
        "呈现唱、白、念、做、打随剧情推进的变化",
        "比较不同剧本的叙事结构差异",
        "总结典型叙事模式及结构特征",
    ])
    scene_df = prepared["scene"]

    # 图1：用可多选场次方块矩阵识别开端、发展、转折、高潮、结局
    render_interactive_stage_tiles(play_id, scene_df)

    # 图2、图3：剧情起伏与表演形式
    c1, c2 = st.columns(2, gap="small")
    with c1:
        plot_scene_curve(play_id, scene_df)
    with c2:
        plot_performance_stack(play_id, scene_df)

    row = fused[fused["play_id"] == str(play_id)]
    if not row.empty:
        r = row.iloc[0]
        cols = st.columns(4)
        metrics = [
            ("叙事模式", r.get("narrative_pattern", "未知")),
            ("平均剧情强度", fmt_num(r.get("intensity_score_mean", np.nan))),
            ("平均节奏指数", fmt_num(r.get("rhythm_score_mean", np.nan))),
            ("高潮位置", fmt_num(r.get("climax_position", np.nan))),
        ]
        for col, (label, value) in zip(cols, metrics):
            with col:
                metric_card(label, value)

    render_question4_insights(fused, play_id)

    # 图4、图5：典型模式与类型结构
    c3, c4 = st.columns([0.9, 1.1], gap="small")
    with c3:
        plot_narrative_pattern_frequency(fused)
    with c4:
        plot_narrative_type_structure(fused)

    section_title("两剧叙事节奏与表演形式对比")
    id_a, id_b, title_a, title_b = two_play_selectors(fused, "q4_compare")
    with st.expander("展开两剧详细对比图", expanded=True):
        a, b = st.columns(2, gap="small")
        with a:
            plot_narrative_pair_curve(scene_df, id_a, id_b, title_a, title_b)
        with b:
            plot_performance_pair(scene_df, id_a, id_b, title_a, title_b)

    render_highlight_search(
        {"场次文本与叙事指标": scene_df, "剧本叙事模式": prepared["narrative"]},
        "q4_search",
        "问题四标红查询｜检索关键阶段、场次文本、表演标记、节奏与叙事模式",
    )


def render_question5(fused: pd.DataFrame, prepared: Dict[str, pd.DataFrame]) -> None:
    task_title(5, "三元关联机制", "关系—主题—叙事协同")
    question_coverage([
        "分析特定角色关系如何承载和推动主题",
        "分析主题结构如何影响叙事策略与剧情组织",
        "分析不同叙事方式如何重塑角色关系",
        "探索典型关联模式与时期协同演化",
        "识别稳定结构特征并支持交互筛选",
    ])

    top_mode = most_common(fused["integrated_mode"])
    top_share = (fused["integrated_mode"].astype(str) == top_mode).mean() if len(fused) else 0
    pair_label, pair_value = strongest_relation_theme_pair(fused)
    top_narr = most_common(fused["narrative_pattern"]) if "narrative_pattern" in fused.columns else "—"

    m1, m2, m3, m4 = st.columns(4, gap="small")
    with m1:
        metric_card("主导协同模式", top_mode)
    with m2:
        metric_card("最高占比", f"{top_share:.1%}")
    with m3:
        metric_card("最强关系—主题", f"{pair_label}（r={pair_value:.2f}）" if pair_label != "—" else "—")
    with m4:
        metric_card("主导叙事模式", top_narr)

    c0, c00 = st.columns([0.82, 1.18], gap="small")
    with c0:
        render_triangle()
    with c00:
        plot_integrated_mode_share(fused)

    c1, c2 = st.columns([1, 1.25], gap="small")
    with c1:
        plot_relation_theme_heatmap(fused)
    with c2:
        plot_sankey(fused)

    c3, c4 = st.columns(2, gap="small")
    with c3:
        plot_narrative_relation_heatmap(fused)
    with c4:
        plot_period_mode_evolution(fused)

    c5, c6 = st.columns(2, gap="small")
    with c5:
        plot_mode_radar(fused)
    with c6:
        plot_mode_structure_heatmap(fused)

    render_highlight_search(
        {
            "融合剧目画像": fused,
            "角色关系边": prepared["edge"],
            "主题结果": prepared["theme"],
            "叙事结果": prepared["narrative"],
        },
        "q5_search",
        "问题五标红查询｜跨关系、主题、叙事与时期结果检索",
    )
def render_data_workshop(
    paths: Dict[str, Optional[Path]],
    raw_data: Dict[str, pd.DataFrame],
    fused: pd.DataFrame,
    prepared: Dict[str, pd.DataFrame],
) -> None:
    section_title("数据工坊｜文件状态、质量检查与结果导出")
    labels = {
        "role": "问题一角色行当",
        "network": "问题二网络指标",
        "edge": "问题二关系边",
        "theme": "问题三主题结果",
        "narrative": "问题四剧本指标",
        "scene": "问题四场次指标",
    }
    cards = st.columns(3)
    for i, (key, label) in enumerate(labels.items()):
        path = paths.get(key)
        color = JADE if path else RED
        with cards[i % 3]:
            st.markdown(
                f"""
<div class="query-box" style="margin-bottom:10px;">
  <div style="font-weight:850;color:{INK};"><span class="status-dot" style="background:{color};"></span>{label}</div>
  <div style="font-size:12px;color:#6e665e;margin-top:6px;">{'已找到：' + html.escape(path.name) if path else '未找到匹配文件'}</div>
  <div style="font-size:13px;color:{RED};margin-top:5px;">记录数：{len(raw_data.get(key, pd.DataFrame()))}</div>
</div>
""",
                unsafe_allow_html=True,
            )

    cols = st.columns(4)
    for col, item in zip(cols, [
        ("融合剧本数", len(fused)),
        ("角色记录", len(prepared["role"])),
        ("关系边记录", len(prepared["edge"])),
        ("场次记录", len(prepared["scene"])),
    ]):
        with col:
            metric_card(*item)

    if not fused.empty:
        st.download_button(
            "下载第五问融合总表",
            data=to_excel_bytes(fused),
            file_name="第五问_关系主题叙事融合总表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    render_highlight_search(
        {
            "角色结果": prepared["role"],
            "关系结果": prepared["edge"],
            "主题结果": prepared["theme"],
            "场次文本": prepared["scene"],
            "融合画像": fused,
        },
        "workshop_search",
        "全库标红查询｜替代大面积原始数据表",
        max_hits=15,
    )

    with st.expander("查看少量关键数据预览（仅前 8 行）", expanded=False):
        preview_source = st.selectbox("选择预览表", list(prepared.keys()), key="preview_source")
        preview_df = prepared.get(preview_source, pd.DataFrame())
        if preview_df.empty:
            st.info("该表为空。")
        else:
            st.dataframe(preview_df.head(8), use_container_width=True, height=270)


# =========================================================
# 8. 主程序
# =========================================================


def main() -> None:
    hero()

    st.sidebar.header("导航")
    page = st.sidebar.radio(
        "选择页面",
        [
            "系统总览",
            "任务一·问题一｜角色行当推断",
            "任务一·问题二｜角色关系网络",
            "任务一·问题三｜核心主题结构",
            "任务一·问题四｜叙事结构分析",
            "任务一·问题五｜三元关联机制",
            "数据工坊",
        ],
        index=0,
    )

    st.sidebar.header("数据路径")
    data_root = st.sidebar.text_input("数据文件夹（云端默认 data）", value=str(DEFAULT_DATA_ROOT))
    paths, raw_data = load_all_data(data_root)
    fused, prepared = build_fused(raw_data)

    if fused.empty:
        st.error("没有成功融合数据。请确认前四问输出文件存在，并且左侧路径填写正确。")
        with st.expander("已检查的文件名"):
            for key, names in FILE_CANDIDATES.items():
                st.write(key, names)
        return

    fused, features = add_pca(fused)
    try:
        fused.to_excel(OUTPUT_DIR / "第五问_关系主题叙事融合总表.xlsx", index=False)
    except Exception:
        pass

    st.sidebar.header("全局筛选")
    play_type_options = sorted(fused["play_type"].dropna().astype(str).unique().tolist())
    period_options = sorted(fused["analysis_period"].dropna().astype(str).unique().tolist())
    selected_types = st.sidebar.multiselect("剧目类型", play_type_options, default=play_type_options)
    selected_periods = st.sidebar.multiselect("历史时期", period_options, default=period_options)

    with st.sidebar.expander("高级筛选", expanded=False):
        theme_options = sorted(fused["theme_pattern"].dropna().astype(str).unique().tolist())
        narrative_options = sorted(fused["narrative_pattern"].dropna().astype(str).unique().tolist())
        mode_options = sorted(fused["integrated_mode"].dropna().astype(str).unique().tolist())
        selected_themes = st.multiselect("主题模式", theme_options, default=theme_options)
        selected_narratives = st.multiselect("叙事模式", narrative_options, default=narrative_options)
        selected_modes = st.multiselect("综合模式", mode_options, default=mode_options)

    filtered = fused[
        fused["play_type"].astype(str).isin(selected_types)
        & fused["analysis_period"].astype(str).isin(selected_periods)
        & fused["theme_pattern"].astype(str).isin(selected_themes)
        & fused["narrative_pattern"].astype(str).isin(selected_narratives)
        & fused["integrated_mode"].astype(str).isin(selected_modes)
    ].copy()

    if filtered.empty:
        st.warning("当前筛选条件下没有数据，请放宽筛选条件。")
        return

    selected_display = st.sidebar.selectbox("当前单剧", filtered["display_name"].tolist(), index=0)
    selected_play_id = selected_display.split("｜")[-1]
    selected_row = fused[fused["play_id"] == selected_play_id].iloc[0]
    selected_title = str(selected_row["play_title"])

    with st.sidebar.expander("数据文件检查", expanded=False):
        for key, path in paths.items():
            st.write(("✅ " if path else "❌ ") + f"{key}：" + (path.name if path else "未找到"))

    if page == "系统总览":
        render_system_overview(filtered, prepared, selected_play_id, selected_title)
    elif page.startswith("任务一·问题一"):
        render_question1(filtered, prepared, selected_play_id)
    elif page.startswith("任务一·问题二"):
        render_question2(filtered, prepared, selected_play_id)
    elif page.startswith("任务一·问题三"):
        render_question3(filtered, prepared, selected_play_id, selected_title)
    elif page.startswith("任务一·问题四"):
        render_question4(filtered, prepared, selected_play_id)
    elif page.startswith("任务一·问题五"):
        render_question5(filtered, prepared)
    elif page == "数据工坊":
        render_data_workshop(paths, raw_data, fused, prepared)


if __name__ == "__main__":
    main()
