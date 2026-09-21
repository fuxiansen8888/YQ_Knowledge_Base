#!/usr/bin/env python3.12
"""把公众号后台导出的单篇明细 .xls 增量并入「公众号文章数据分析表.xlsx」。

raw/data/wechat/ 是新数据的投递口：把导出丢进去、跑一次脚本就行。

增量规则（每次运行）：
    新标题     → 追加一行
    已有标题   → 只刷新数据列（阅读/分享/推荐占比这类客观值），人工填的列一律不动
    行永不删除 → raw/data/wechat/ 里删掉某份导出，表里那一行仍保留，历史不因文件消失而丢
    人工填的列（D 字数 / E 内容方向 / F 人群匹配度 / G 核心价值类型 / P 跳出率）
               按标题原样保留，只有空着的才填默认值（方向与价值类型按标题关键词猜，匹配度按三档阈值算）
    行序      按发布日期倒序（最新在最上面），序号随之重编 1..N

后台导出有两种结构，脚本按表头取列、不认死行号：
    人数版：数据概况用「阅读(人)/分享(人)/收藏(人)」，趋势明细只有 阅读人数/分享人数
    次数版：数据概况用「阅读(次)/分享(次)/在看(次)/点赞(次)」，趋势明细带 阅读人数/阅读次数/分享人数/分享次数
两版都从趋势明细的「全部」渠道行还原人数口径。

「方向×价值类型透视」：按 内容方向×核心价值类型×人群匹配度 聚合「文章数据记录」，自动重写。

用法:
    python3.12 import_wechat_export.py

只要有一份导出解析失败，脚本整体停下、不写文件。
"""
from __future__ import annotations

import glob
import os
import re
import shutil
import zipfile
from copy import copy

from collections import Counter

import openpyxl
import xlrd
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from xml.sax.saxutils import escape

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE, "公众号文章数据分析模板.xlsx")
OUTPUT = os.path.join(BASE, "公众号文章数据分析表.xlsx")
# 原始导出属于只读素材层，与脚本产物分开放
DATA_DIR = os.path.normpath(os.path.join(BASE, "..", "..", "raw", "data", "wechat"))
SHEET = "文章数据记录"
PIVOT_SHEET = "方向×价值类型透视"
REGIME_SHEET = "流量结构透视"
MONTH_SHEET = "月度趋势"

FIRST_ROW, LAST_ROW = 6, 49  # 数据区行范围（装满了脚本会自己往下长，见 ensure_capacity）
SAMPLE_ROWS = range(6, 11)  # 模板自带的示例行
DATA_STYLE_ROW = 30  # 模板里本来就是数据行的一行，迁移时抄它的样式
NOTES_FIRST, NOTES_ROWS = LAST_ROW + 1, 9  # 填写说明整块的位置（模板原在第 32~39 行，脚本往下挪）
GROWTH_HEADROOM = 10  # 数据区不够高时，一次多留几行

# 分区标题：用于给重名标签（如「未知」）划清归属
SECTIONS = {"数据概况", "阅读转化", "性别分布", "年龄分布", "地域分布"}
# 阅读渠道（趋势明细的「传播渠道」列）：先私域后公域，最后留「其他」兜底
CHANNELS = ("公众号消息", "公众号主页", "聊天会话", "朋友圈", "朋友在看", "推荐", "搜一搜", "其他")
TREND_METRICS = ("阅读人数", "阅读次数", "分享人数", "分享次数")
ALL_CHANNEL = "全部"

# 分类规则，按顺序生效、命中即停：
#   ① TITLE_RULES  —— 逐篇过目后指定的结论（标题里含左侧关键词即命中），最可靠
#   ② 关键词表      —— 新文章兜底推断；注意别用会误命中的裸词（「花」会撞「花钱」，「如何」会撞「如何感受美」）
#   ③ 兜底值
TITLE_RULES = (
    # (标题关键词, 内容方向, 核心价值类型)
    ("越提醒，孩子越磨蹭", "育儿育己", "认知启发"),
    ("女儿站在滑梯边", "育儿育己", "情绪共鸣"),
    ("不肯分享玩具", "育儿育己", "认知启发"),
    ("梳妆台前清晨7点", "生活感悟", "认知启发"),
    ("更看重精神自由", "生活感悟", "情绪共鸣"),
    ("好友单程1.5小时", "生活感悟", "情绪共鸣"),
    ("我救自己万万次", "生活感悟", "认知启发"),
    ("接放学没跑第一", "育儿育己", "认知启发"),
    ("凌晨2点的夫妻夜谈", "生活感悟", "认知启发"),
    ("越来越讨厌自己的", "生活感悟", "认知启发"),
    ("为了在北方闻一口花香", "植物美学", "审美滋养"),
    ("太好了，我才35岁", "生活感悟", "情绪共鸣"),
    ("LV诉奶茶店", "生活感悟", "实用方法"),
    ("倍速时代", "追剧观影", "认知启发"),
    ("送女儿上幼儿园第一天", "育儿育己", "情绪共鸣"),
    ("你看姐姐多勇敢", "育儿育己", "认知启发"),
    ("谁规定人必须全年开花", "植物美学", "认知启发"),
    ("姐姐是喜欢你呀", "育儿育己", "认知启发"),
    ("连夜逃去广州", "育儿育己", "情绪共鸣"),
    ("自我和解说明书", "生活感悟", "情绪共鸣"),
    ("一株芦荻的启示", "植物美学", "审美滋养"),
    ("生活不在别处", "生活感悟", "情绪共鸣"),
    ("破晓东方", "追剧观影", "认知启发"),
    ("生而和平", "深度思考", "认知启发"),
    ("阅读的终极目的", "生活感悟", "认知启发"),
    ("高敏感者的自我救赎", "生活感悟", "认知启发"),
    ("沉默的荣耀", "追剧观影", "认知启发"),
    ("菊花羞作倾城色", "植物美学", "审美滋养"),
    ("一切诸相，皆是虚妄", "深度思考", "认知启发"),
    ("所有人都将渐行渐远", "生活感悟", "情绪共鸣"),
    ("路边的柏香里", "植物美学", "审美滋养"),
    ("恶女花魁", "追剧观影", "审美滋养"),
    ("人生四千周", "生活感悟", "认知启发"),
    ("从小镇批发市场看人生", "生活感悟", "认知启发"),
    ("专注力UP", "育儿育己", "实用方法"),
    ("婆婆教我的事", "生活感悟", "认知启发"),
    ("柳如是", "追剧观影", "审美滋养"),
    ("我永远不会回到以前的生活", "育儿育己", "认知启发"),
    ("草木记忆", "植物美学", "审美滋养"),
    ("蝉的信仰", "生活感悟", "认知启发"),
)

# 「内容方向」关键词表：按顺序匹配，先命中者优先
DIRECTION_RULES = (
    (("电影", "影视", "影院", "追剧", "观影", "剧集", "电视剧", "看剧", "重看", "剧透", "小评"),
     "追剧观影"),
    (("植物", "养花", "插花", "花香", "花卉", "花园", "花市", "绿植", "园艺", "盆栽",
      "阳台", "芦荻", "菊花", "落叶", "开花"), "植物美学"),
    (("孩子", "女儿", "儿子", "育儿", "妈妈", "亲子", "家长", "宝宝", "教育", "幼儿园", "黏人精"), "育儿育己"),
    (("战略", "战争", "智慧", "哲学", "虚妄", "信仰", "文明", "认知", "本质"), "深度思考"),
    (("生活", "婚姻", "夫妻", "社交", "自己", "人生", "成长", "日常", "拒绝", "生日"), "生活感悟"),
)
DIRECTION_FALLBACK = "生活感悟"

# 「核心价值类型」关键词表，同上
VALUE_RULES = (
    (("方法", "步骤", "清单", "技巧", "指南", "行动", "个改变", "怎么做"), "实用方法"),
    (("本质", "认知", "为什么", "原理", "机制", "真相", "思考", "觉醒", "重构", "启发", "启示"), "认知启发"),
    (("情绪", "委屈", "难过", "心疼", "共情", "焦虑", "眼泪", "释怀", "救赎", "孤独", "勇气", "原谅"), "情绪共鸣"),
    (("审美", "美", "治愈", "诗意", "意境"), "审美滋养"),
)
VALUE_FALLBACK = "纯粹记录"

# 标题里看不出价值取向时的二级兜底：按内容方向推最可能的类型，避免全堆到「纯粹记录」
VALUE_BY_DIRECTION = {
    "育儿育己": "情绪共鸣",
    "植物美学": "审美滋养",
    "追剧观影": "认知启发",
    "深度思考": "认知启发",
    "生活感悟": "情绪共鸣",
}

# 人群匹配度三档：都达标=高匹配，一项达标=中匹配，都不达标=低匹配
FEMALE_THRESHOLD, YOUNG_THRESHOLD = 0.6, 0.5
MATCH_TEXT = ("高匹配", "中匹配", "低匹配")

# 表结构：五层 + 每层列名。列位置由这张表决定，写表头、取列、验列序都以它为准
LAYERS = (
    ("第一层：基础信息", ("序号", "发布日期", "标题", "字数", "内容方向", "人群匹配度", "核心价值类型")),
    ("第二层：流量表现", ("阅读人数", "目标人群阅读量", "送达人数", "公众号消息阅读人数", "首次分享人数",
                          "订阅打开率", "订阅分享率")),
    ("第三层：内容质量", ("完读率", "平均停留时长", "微信收藏人数", "收藏率")),
    ("第四层：传播转化", ("点赞数", "在看数", "分享人数", "分享率",
                           "分享产生的阅读人数", "分享扩散系数", "读后关注人数")),
    ("第五层：人群匹配", ("女性读者占比", "25-45岁读者占比", "46岁以上占比", "目标人群匹配度",
                           "单篇涨粉数", "涨粉转化率")),
    ("第六层：阅读渠道构成", tuple(f"{name}占比" for name in CHANNELS)),
)
# 列的字母位置：序号=A、阅读人数=I、完读率=O… 全部由 LAYERS 推出来，避免手写字母写错
COLUMNS = {name: get_column_letter(index)
           for index, name in enumerate((n for _, names in LAYERS for n in names), start=1)}
LAST_COLUMN = len(COLUMNS)  # AC 列（29 列）
DATA_COLUMNS = tuple(get_column_letter(i) for i in range(2, LAST_COLUMN + 1))  # B~AC，快照整行用
MANUAL_NAMES = ("字数",)  # 导出文件没有、只能人工填的列
JUDGMENT_NAMES = ("内容方向", "人群匹配度", "核心价值类型")  # 判断列：只在空着时填默认值
NEVER_WRITTEN = tuple(COLUMNS[name] for name in MANUAL_NAMES)  # 脚本从不写入的列
JUDGMENT_COLS = tuple(COLUMNS[name] for name in JUDGMENT_NAMES)
SAMPLE_CLEAR_COLS = NEVER_WRITTEN + JUDGMENT_COLS  # 首次运行时清掉模板示例行的判断/人工列
CARRY_COLS = NEVER_WRITTEN + JUDGMENT_COLS  # 只有这些列从历史行继承；脚本自己的列一律以本次导出为准

# 百分比列显示格式。订阅打开率/搜一搜占比数值很小，给两位小数才看得出来；分享扩散系数是倍数，不是百分比
PERCENT_FORMATS = {COLUMNS[name]: "0.0%" for name in
                   ("订阅分享率", "推荐占比", "完读率", "收藏率", "分享率",
                    "女性读者占比", "25-45岁读者占比", "涨粉转化率")}
PERCENT_FORMATS[COLUMNS["订阅打开率"]] = "0.00%"
PERCENT_FORMATS[COLUMNS["搜一搜占比"]] = "0.00%"
PERCENT_FORMATS[COLUMNS["分享扩散系数"]] = "0.00"
PERCENT_FORMATS[COLUMNS["46岁以上占比"]] = "0.0%"
for _channel in CHANNELS:  # 渠道占比：小的只有 0.0x%，给两位小数
    PERCENT_FORMATS[COLUMNS[f"{_channel}占比"]] = "0.00%"

# 公式列：把口径写在表里，改了源数据会自动重算。除数为空或为 0 都要挡掉，否则 Excel 报 #DIV/0!
# 数值型公式的分子/分母：既用于生成公式，也用于算缓存值（见 inject_cached_results）
FORMULA_INPUTS = {
    "订阅打开率": ("公众号消息阅读人数", "送达人数"),
    "订阅分享率": ("首次分享人数", "公众号消息阅读人数"),
    "分享扩散系数": ("分享产生的阅读人数", "分享人数"),  # 分母用第四层的分享人数（它与已删的「总分享人数」同源）
    "收藏率": ("微信收藏人数", "阅读人数"),
    "分享率": ("分享人数", "阅读人数"),
    "涨粉转化率": ("单篇涨粉数", "阅读人数"),
}

def ratio_formula(numerator: str, denominator: str) -> str:
    """按分子/分母列生成带保护的除法公式（0 与空值都挡掉，否则 Excel 报 #DIV/0!）。"""
    top, bottom = COLUMNS[numerator], COLUMNS[denominator]
    return (f'=IF(OR({bottom}{{r}}="",{bottom}{{r}}=0,{top}{{r}}=""),"",{top}{{r}}/{bottom}{{r}})')


FORMULAS = {
    **{COLUMNS[name]: ratio_formula(*pair) for name, pair in FORMULA_INPUTS.items()},
    # 目标人群匹配度是文本公式：女性≥60% 且 25-45岁≥50% 为高匹配，一项达标为中匹配，都不达标为低匹配
    COLUMNS["目标人群匹配度"]: (
        f'=IF(OR({COLUMNS["女性读者占比"]}{{r}}="",{COLUMNS["25-45岁读者占比"]}{{r}}=""),"",'
        f'IF(AND({COLUMNS["女性读者占比"]}{{r}}>=0.6,{COLUMNS["25-45岁读者占比"]}{{r}}>=0.5),"高匹配",'
        f'IF(OR({COLUMNS["女性读者占比"]}{{r}}>=0.6,{COLUMNS["25-45岁读者占比"]}{{r}}>=0.5),"中匹配","低匹配")))'),
}


PIVOT_FIRST_ROW, PIVOT_LAST_ROW = 5, 27  # 透视表数据区行范围
METRICS = ("订阅打开率", "完读率", "分享率", "女性占比", "25-45岁占比", "涨粉转化率")  # 对应透视表 E~J 列
EXTRA_METRICS = ("推荐占比",)  # 只用于诊断结论，不占透视表列
UNFILLED = "（未填写）"

# 数据区的统一外观：一律微软雅黑 10、水平垂直都居中。模板里个别格子会沿用默认宋体 11、不居中，看着花
DATA_FONT, DATA_SIZE = "微软雅黑", 10
DATA_ROW_HEIGHT, PIVOT_ROW_HEIGHT = 24, 104  # 透视表行更高：结论是三段多行文字（v2 最长占 7 行）
PIVOT_WRAP = True  # 透视表「结论与建议」是长文本，要自动换行
CONCLUSION_WIDTH = 62  # 透视表 K 列宽度（要放下定位＋诊断＋动作）
PIVOT_LEFT_COLS = ("M",)  # 结论是成段文字，居左更好读；其余数值列仍居中

# 第一个表里「越高越好」的关键指标列，加红-黄-绿色阶，好一眼看出谁高谁低于全表均值。
# 透视表的诊断结论就是拿这几项跟全表均值比的，色阶是它的可视化对照。
COLOR_SCALE_COLS = tuple(COLUMNS[name] for name in
                         ("订阅打开率", "完读率", "分享率", "女性读者占比", "25-45岁读者占比", "涨粉转化率"))
COLOR_SCALE_RULE = dict(start_type="min", start_color="FFF8696B",
                        mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
                        end_type="max", end_color="FF63BE7B")
# 第六层：按行染色——每行的 8 个渠道互相比较，一眼看出这篇的主力渠道。
# 用单色顺序色阶（浅→深）而不是红-黄-绿：渠道占比高不代表好，红绿会给出错误的褒贬暗示。
CHANNEL_ROW_RULE = dict(start_type="min", start_color="FFFFFFFF",
                        mid_type="percentile", mid_value=50, mid_color="FFD9E1F2",
                        end_type="max", end_color="FF4472C4")

# 46岁以上占比是**反向**指标（越低越好），所以色阶要反过来：低=绿、高=红
REVERSE_RULES = {COLUMNS["46岁以上占比"]: dict(start_type="min", start_color="FF63BE7B",
                                               mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
                                               end_type="max", end_color="FFF8696B")}

# 阅读人数离群值太大（中位 116、最大 38596），线性色阶会让其余 37 篇全挤到红端，改按分位数
OUTLIER_RULES = {COLUMNS["阅读人数"]: dict(start_type="percentile", start_value=10, start_color="FFF8696B",
                                           mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
                                           end_type="percentile", end_value=90, end_color="FF63BE7B"),
                  COLUMNS["目标人群阅读量"]: dict(start_type="percentile", start_value=10,
                                                 start_color="FFF8696B", mid_type="percentile", mid_value=50,
                                                 mid_color="FFFFEB84", end_type="percentile", end_value=90,
                                                 end_color="FF63BE7B")}


def norm(label: object) -> str:
    """标签归一化：统一中英文括号、去掉空白，避免「阅读(人)」与「阅读（人）」不匹配。"""
    return re.sub(r"[（）()\s]", "", str(label))


def num(value: object) -> float:
    """'69.42%' → 0.6942；'0.099174' → 0.099174；121.0 → 121.0。"""
    text = str(value).strip()
    return float(text.rstrip("%")) / 100 if text.endswith("%") else float(text)


def numeric_like(value: object) -> bool:
    """像数字就当数字处理：'69.42%' / 0.6942 / 51 都算，'人数' / '占比' 这种表头不算。"""
    text = str(value).strip()
    if text.endswith("%"):
        return True
    try:
        float(text)
        return True
    except ValueError:
        return False


def distribution(raw: dict) -> dict:
    """把一组分布值统一成占比。

    两种导出的写法不同：人数版给百分比（'69.42%'），次数版给**绝对人数**（女 51 / 男 13）。
    按合计判断：合计 >1.5 说明是人数或 100 制百分数，除以合计归一到 1。
    """
    numbers = {label: num(value) for label, value in raw.items() if numeric_like(value)}
    total = sum(numbers.values())
    if total <= 0:
        return {}
    scale = total if total > 1.5 else 1.0
    return {label: value / scale for label, value in numbers.items()}


def is_elder_age(label: str) -> bool:
    """两种导出的年龄段分档名不同：46-55/56-65/65岁以上（人数版）、46-60/60岁以上（次数版）。"""
    return label.startswith(("46-", "56-")) or label in ("65岁以上", "60岁以上")


def share(value: object) -> float:
    """占比归一化：'69.42%' → 0.6942；数字 70.0 → 0.70；0.6942 → 0.6942。"""
    number = num(value)
    return number / 100 if number > 1 else number


def ratio(numerator: object, denominator: object) -> float | None:
    """能算才算：任一非数字或分母为 0 时返回 None。"""
    if not isinstance(numerator, (int, float)) or not isinstance(denominator, (int, float)):
        return None
    return numerator / denominator if denominator else None


def close(a: object, b: object) -> bool:
    """数值比较带容差：xlsx 只存 15 位有效数字，直接比会报出一堆假变化。"""
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) < 1e-9
    return a == b


def stored(value: object) -> object:
    """写表用的精度（6 位小数），也让「变化检测」和写入值口径一致。"""
    return round(value, 6) if isinstance(value, float) else value


def mean(values: list) -> float | None:
    return sum(values) / len(values) if values else None


def median(values: list) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    middle = len(ordered) // 2
    return ordered[middle] if len(ordered) % 2 else (ordered[middle - 1] + ordered[middle]) / 2


def pct(value: float | None) -> str:
    return "—" if value is None else f"{value:.1%}"


def match_level(female: object, young: object) -> str | None:
    """人群匹配度三档，阈值与 FORMULAS['Z'] 保持一致。"""
    if not isinstance(female, (int, float)) or not isinstance(young, (int, float)):
        return None
    hits = int(female >= FEMALE_THRESHOLD) + int(young >= YOUNG_THRESHOLD)
    return MATCH_TEXT[2 - hits]


def keyword_match(title: str, rules: tuple, fallback: str) -> str:
    for words, label in rules:
        if any(word in title for word in words):
            return label
    return fallback


def classify(title: str) -> tuple[str, str]:
    """给一篇文章定 (内容方向, 核心价值类型)：逐篇规则 → 关键词 → 按方向二级兜底 → 兜底值。"""
    for marker, direction, value in TITLE_RULES:
        if marker in title:
            return direction, value
    direction = keyword_match(title, DIRECTION_RULES, DIRECTION_FALLBACK)
    value = keyword_match(title, VALUE_RULES, "")
    return direction, value or VALUE_BY_DIRECTION.get(direction, VALUE_FALLBACK)


def default_keys(title: str) -> dict[str, str]:
    """第一层两个判断列的默认值：内容方向、核心价值类型。"""
    direction, value = classify(title)
    return {"E": direction, "G": value}


def read_export(path: str) -> dict:
    """解析一份导出文件，返回 {标题, 发布日期, 指标, 告警}。拿不到的指标不出现在结果里。"""
    sheet = xlrd.open_workbook(path).sheet_by_index(0)
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]
    title = str(rows[0][1]).strip()

    # 趋势明细：按表头行（含「日期」「传播渠道」）定位，再按列名取值，两版结构通吃
    head = next((i for i, row in enumerate(rows)
                 if any(str(v).strip() == "日期" for v in row)
                 and any(str(v).strip() == "传播渠道" for v in row)), None)
    if head is None:
        raise ValueError("找不到趋势明细表头（日期/传播渠道）")
    columns = {str(v).strip(): j for j, v in enumerate(rows[head])}
    trend: list[dict] = []
    for row in rows[head + 1:]:
        date = str(row[columns["日期"]]).strip() if len(row) > columns["日期"] else ""
        channel = str(row[columns["传播渠道"]]).strip() if len(row) > columns["传播渠道"] else ""
        if not date or not channel:
            break
        item = {"日期": date, "渠道": channel}
        for name in TREND_METRICS:
            column = columns.get(name)
            if column is not None and len(row) > column and isinstance(row[column], float):
                item[name] = int(row[column])
        trend.append(item)
    if not trend:
        raise ValueError("趋势明细没有数据行")
    totals = [item for item in trend if item["渠道"] == ALL_CHANNEL]
    if not totals:
        raise ValueError("趋势明细里没有「全部」渠道行")

    # 其余分区按「分区 + 指标」取值，趋势块整段跳过
    labels: dict = {}
    section: str | None = None
    for index, row in enumerate(rows):
        if head <= index <= head + len(trend):
            continue
        key = str(row[1]).strip() if len(row) > 1 else ""
        if key in SECTIONS:
            section = key
            continue
        if not key or len(row) < 3 or row[2] == "":
            continue
        labels[(section, norm(key))] = row[2]

    def pick(name_section: str, *names: str):
        for name in names:
            if (name_section, norm(name)) in labels:
                return labels[(name_section, norm(name))]
        return None

    def trend_sum(name: str):
        values = [item[name] for item in totals if name in item]
        return sum(values) if len(values) == len(totals) else None

    warnings: list[str] = []

    def required(raw, name: str):
        if raw is None:
            raise ValueError(f"缺少指标: {name}")
        return num(raw)

    readers = trend_sum("阅读人数")
    if readers is None:
        raise ValueError("趋势明细缺少「阅读人数」列")
    stated_readers = pick("数据概况", "阅读(人)")
    if stated_readers is not None and num(stated_readers) != readers:
        raise ValueError(f"趋势明细阅读人数合计 {readers} 与数据概况「阅读(人)」{stated_readers} 不一致")
    shares = trend_sum("分享人数")
    stated_shares = pick("数据概况", "分享(人)")
    if shares is None and stated_shares is None:
        raise ValueError("既没有分享人数列也没有「分享(人)」")
    if shares is not None and stated_shares is not None and num(stated_shares) != shares:
        raise ValueError(f"趋势明细分享人数合计 {shares} 与数据概况「分享(人)」{stated_shares} 不一致")
    if shares is None:
        shares = int(num(stated_shares))

    # 阅读转化分区的原始指标：人数版是「人」，次数版是「次」（口径混用，用哪版就按哪版算）
    def conversion(*names: str):
        raw = pick("阅读转化", *names)
        return num(raw) if raw is not None else None

    delivered = int(required(pick("阅读转化", "送达人数"), "送达人数"))
    msg_reads = conversion("公众号消息阅读人数", "公众号消息阅读次数")
    first_shares = conversion("首次分享人数", "首次分享次数")
    spread_reads = conversion("分享产生的阅读人数", "分享产生的阅读次数")
    follows = int(required(pick("数据概况", "新增关注（人）", "阅读后关注（人）"), "新增关注/阅读后关注"))

    values = {
        "B": min(item["日期"] for item in totals),  # 发布日期取趋势明细首日
        "C": title,
        COLUMNS["阅读人数"]: readers,
        COLUMNS["送达人数"]: delivered,
        COLUMNS["公众号消息阅读人数"]: int(msg_reads) if msg_reads is not None else None,
        COLUMNS["首次分享人数"]: int(first_shares) if first_shares is not None else None,
        COLUMNS["分享产生的阅读人数"]: int(spread_reads) if spread_reads is not None else None,
        COLUMNS["完读率"]: share(required(pick("数据概况", "完读率"), "完读率")),
        COLUMNS["平均停留时长"]: int(required(pick("数据概况", "平均停留时长(秒)"), "平均停留时长(秒)")),
        COLUMNS["点赞数"]: int(required(pick("数据概况", "点赞(人)", "点赞(次)"), "点赞")),
        COLUMNS["在看数"]: int(required(pick("数据概况", "在看(人)", "在看(次)"), "在看")),
        COLUMNS["分享人数"]: shares,
        COLUMNS["读后关注人数"]: follows,
        COLUMNS["单篇涨粉数"]: follows,
    }
    # 可选列：拿得到就写，拿不到就不写（保留人工填的值）
    optional = {
        COLUMNS["微信收藏人数"]: pick("数据概况", "收藏(人)"),  # 次数版没有收藏
    }
    sex_dist = distribution({label: value for (section, label), value in labels.items()
                             if section == "性别分布"})
    age_dist = distribution({label: value for (section, label), value in labels.items()
                             if section == "年龄分布"})
    if "女" in sex_dist:
        optional[COLUMNS["女性读者占比"]] = sex_dist["女"]
    young = [age_dist[label] for label in ("26-35岁", "36-45岁") if label in age_dist]
    if young:
        optional[COLUMNS["25-45岁读者占比"]] = sum(young)
    elder = [value for label, value in age_dist.items() if is_elder_age(label)]
    if elder:  # 46 岁以上＝老龄化反向指标
        optional[COLUMNS["46岁以上占比"]] = sum(elder)
    for col, raw in optional.items():
        if raw is not None:
            values[col] = int(num(raw)) if col == COLUMNS["微信收藏人数"] else share(raw)
    for channel in CHANNELS:  # 第六层：把趋势明细里各渠道（按天多行）汇总，再除以阅读人数得占比
        col = COLUMNS[f"{channel}占比"]
        channel_share = ratio(sum(item["阅读人数"] for item in trend if item["渠道"] == channel), readers)
        if channel_share is None:
            continue
        if channel_share > 1:  # 渠道去重差可能让单渠道超过总量，宁可不写
            warnings.append(f"{channel}占比 {pct(channel_share)} > 100%，已跳过 {col} 列")
            continue
        values[col] = channel_share  # 该渠道没流量就是 0，也要落成 0% 而不是留空

    for name, col in (("完读率", "完读率"), ("推荐占比", "推荐占比"), ("搜一搜占比", "搜一搜占比"),
                      ("女性读者占比", "女性读者占比"), ("25-45岁读者占比", "25-45岁读者占比"),
                      ("46岁以上占比", "46岁以上占比")):
        value = values.get(COLUMNS[col])
        if value is not None and not 0 <= value <= 1:
            raise ValueError(f"{name} 越界: {value}")
    # 订阅打开率/订阅分享率是「人（或次）/人」的比值：次数版混用口径时可能 >100%，如实写但提示
    if values.get(COLUMNS["25-45岁读者占比"]) is not None:  # 成绩单口径：阅读量中落到目标人群的部分
        values[COLUMNS["目标人群阅读量"]] = round(readers * values[COLUMNS["25-45岁读者占比"]])
    for name in ("订阅打开率", "订阅分享率"):
        value = values.get(COLUMNS[name])
        if value is not None and not 0 <= value <= 1:
            warnings.append(f"{name} {pct(value)} 超出 0~100%（次数版按「次」计算所致），已如实写入")
    return {"file": os.path.basename(path), "values": values, "warnings": warnings}


def is_article_row(sheet, row: int) -> bool:
    """已导入的文章行 = 有标题且有阅读人数；模板里的零散字符（如 C8 的 "s"）不算。"""
    return bool(sheet.cell(row, 3).value) and isinstance(sheet[f'{COLUMNS["阅读人数"]}{row}'].value, (int, float))


def find_notes_row(sheet) -> int | None:
    for row in range(FIRST_ROW + 5, sheet.max_row + 1):
        if str(sheet[f"A{row}"].value or "").strip() == "【填写说明】":
            return row
    return None


def migrate_layout(sheet) -> bool:
    """把【填写说明】整块挪到 NOTES_FIRST 行，腾出来的行改成数据行样式，数据区随 FIRST/LAST_ROW 扩大。"""
    current = find_notes_row(sheet)
    if current is None:
        raise SystemExit("找不到【填写说明】行，表结构不是预期版本")
    moved = current != NOTES_FIRST
    if moved:
        offset = NOTES_FIRST - current
        rows = list(range(current, current + NOTES_ROWS))
        merges = {str(r) for r in sheet.merged_cells.ranges}
        for row in rows:  # 先整块拆掉合并，否则往里写会撞上只读的 MergedCell
            merged = f"A{row}:AB{row}"
            if merged in merges:
                sheet.unmerge_cells(merged)
        # 目标位置可能和原位重叠，先快照到内存再落笔
        snapshot = [[(sheet.cell(row, column).value, copy(sheet.cell(row, column)._style))
                     for column in range(1, 29)] for row in rows]
        heights = [sheet.row_dimensions[row].height for row in rows]
        for index, row in enumerate(rows):
            for column, (value, style) in enumerate(snapshot[index], start=1):
                target = sheet.cell(row + offset, column)
                target.value = value
                target._style = style
            sheet.merge_cells(f"A{row + offset}:AB{row + offset}")
            sheet.row_dimensions[row + offset].height = heights[index]
        for row in rows:  # 原位腾出来的行清空（和新块重叠的除外，它们已被新内容覆盖）
            if not NOTES_FIRST <= row < NOTES_FIRST + NOTES_ROWS:
                for column in range(1, 29):
                    sheet.cell(row, column).value = None

    for row in range(FIRST_ROW, LAST_ROW + 1):  # 数据行样式（腾空的行和新增的行都要补）
        for column in range(1, 29):
            cell = sheet.cell(row, column)
            if cell.value in (None, ""):
                cell._style = copy(sheet.cell(DATA_STYLE_ROW, column)._style)
        sheet.row_dimensions[row].height = sheet.row_dimensions[DATA_STYLE_ROW].height
    return moved


def write_layout(sheet) -> bool:
    """按 LAYERS 重铺第 4 行分组横幅与第 5 行列名。

    列口径改过（打开率、阅读次数已删，新增订阅打开率/订阅分享率/分享扩散系数），模板原来的表头
    已经对不上。所以每次运行都按 LAYERS 重写表头两行 + 收拢横幅合并区，改了 LAYERS 就自动生效。
    """
    before = [(sheet.cell(4, c).value, sheet.cell(5, c).value) for c in range(1, LAST_COLUMN + 1)]
    old_headers = {get_column_letter(c): sheet.cell(5, c).value for c in range(1, LAST_COLUMN + 1)}
    # 人工列会随口径改版换位置（跳出率：旧结构在 P，新结构在 Q），先按语义搬过去，
    # 否则快照按列字母搬运会把旧位置上的别的列的值串进来。
    for name in MANUAL_NAMES:
        was = next((letter for letter, header in old_headers.items() if header == name), None)
        now = COLUMNS[name]
        if was and was != now:
            for row in range(FIRST_ROW, LAST_ROW + 1):
                sheet[f"{now}{row}"] = sheet[f"{was}{row}"].value
                sheet[f"{was}{row}"] = None
    for letter, header in old_headers.items():  # 旧布局里有、新布局里没有的列（如已删的跳出率）：数据清掉
        if header and header not in COLUMNS:
            for row in range(FIRST_ROW, LAST_ROW + 1):
                sheet[f"{letter}{row}"] = None
    for rng in list(sheet.merged_cells.ranges):  # 拆掉旧横幅（第 1/2 行标题、说明区不动）
        if rng.min_row == rng.max_row == 4:
            sheet.unmerge_cells(str(rng))
    for name in ("A1", "A2"):  # 标题/副标题横幅延伸到最后一列
        for rng in list(sheet.merged_cells.ranges):
            if str(rng).startswith(f"{name}:"):
                sheet.unmerge_cells(str(rng))
        sheet.merge_cells(f"{name}:{get_column_letter(LAST_COLUMN)}{name[1:]}")
    for rng in list(sheet.merged_cells.ranges):  # 填写说明行也延伸到最后一列
        if rng.min_row >= NOTES_FIRST:
            sheet.unmerge_cells(str(rng))
            sheet.merge_cells(f"A{rng.min_row}:{get_column_letter(LAST_COLUMN)}{rng.min_row}")
    style_notes(sheet)

    banner_style = copy(sheet.cell(4, 1)._style)
    header_style = copy(sheet.cell(5, 1)._style)
    column = 1
    for banner, names in LAYERS:
        first = column
        for name in names:
            sheet.cell(5, column).value = name
            sheet.cell(5, column)._style = header_style
            column += 1
        sheet.cell(4, first).value = banner
        for c in range(first, column):
            sheet.cell(4, c)._style = banner_style
        sheet.merge_cells(f"{get_column_letter(first)}4:{get_column_letter(column - 1)}4")
    for c in range(column, LAST_COLUMN + 6):  # 清掉超出新结构的旧表头
        sheet.cell(4, c).value = None
        sheet.cell(5, c).value = None
    for c in range(sheet.max_column, LAST_COLUMN, -1):  # 表变窄时（删列）整列删掉，别留空框
        sheet.delete_cols(c)
    after = [(sheet.cell(4, c).value, sheet.cell(5, c).value) for c in range(1, LAST_COLUMN + 1)]
    return before != after


def document_basis(sheet) -> bool:
    """让透视表的文字跟当前口径一致：列名按 METRICS 走，副标题与判定说明写清基准。

    模板原话还写着「需手动填入汇总数据」，第一列还叫「平均打开率」（那个口径已删）。
    """
    changed = False
    for col, name in zip("EFGHIJ", METRICS):
        label = f"平均{name}"
        if sheet[f"{col}4"].value != label:
            sheet[f"{col}4"] = label
            changed = True
    if "全表均值" in str(sheet["A2"].value or ""):
        sheet["A2"] = ("按 内容方向×价值类型×匹配度 交叉分析；结论按流量结构分层比较"
                       "（算法/混合/私域各自的同层水平，详见「流量结构透视」），样本<3 篇仅供参考")
        changed = True
    if str(sheet["A31"].value or "").startswith(("结论判断标准", "赛道判定")):  # 只认旧文案，新文案不再重写
        sheet["A31"] = ("结论 v2：先看流量结构（算法/混合/私域），所有比较都在**同层内**进行——"
                        "推荐流量是冷流量，完读率与分享率天然低于私域，跨层比会误判；"
                        "定位给轨道（增长轨/基本盘轨）与最弱维度，动作按「轨道×短板」给")
        changed = True
    return changed


CHANNEL_NOTE = ("第六层（阅读渠道构成）：各渠道占比 = 该渠道阅读人数 ÷ 阅读人数"
                "（趋势明细里按天多行，先汇总再算），自动计算；渠道去重差可能让各项之和不等于 100%")

NOTE_UPDATES = (
    ("第六层（阅读渠道构成）", CHANNEL_NOTE),
    ("第二层（流量表现）",
     "第二层（流量表现）：阅读人数/送达人数/公众号消息阅读人数/首次分享人数从后台导出 | "
     "目标人群阅读量＝阅读人数×25-45岁占比（对目标读者的触达，成绩单口径）| "
     "订阅打开率=公众号消息阅读人数÷送达人数、订阅分享率=首次分享人数÷公众号消息阅读人数（自动算）| "
     "推荐占比/搜一搜占比从单篇详情查看"),
    ("第三层（内容质量）", "第三层（内容质量）：完读率/停留时长从单篇详情查看 | 收藏率自动计算"),
    ("核心指标参考",
     "核心指标参考：分享率≥2%合格≥10%强传播 | 完读率≥40%有推荐≥50%进优质池（订阅打开率的合格线请按后台同类账号自定）"),
)


def ensure_channel_note(sheet) -> bool:
    """把第六层说明插到第五层之后（老文件里只有到第五层的说明）。"""
    rows = list(range(NOTES_FIRST, NOTES_FIRST + NOTES_ROWS))
    texts = [str(sheet[f"A{r}"].value or "") for r in rows]
    if any(text.startswith("第六层") for text in texts):
        return False
    five = next((i for i, text in enumerate(texts) if text.startswith("第五层")), None)
    if five is None or len(rows) <= five + 2:
        return False
    insert_at = five + 1
    for offset, text in enumerate(texts[insert_at:-1]):  # 原有的后续说明整体下移一行
        sheet[f"A{rows[insert_at + 1 + offset]}"] = text or None
    sheet[f"A{rows[insert_at]}"] = CHANNEL_NOTE
    return True


NOTE_HEAD_FONT = Font(name="微软雅黑", size=11, bold=True, color="FF1C3A5E")
NOTE_BODY_FONT = Font(name="微软雅黑", size=9, italic=True, color="FF8A94A6")
NOTE_BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_notes(sheet) -> int:
    """说明区样式归位：首行 11 号加粗深蓝，其余 9 号斜体灰字左对齐（模板本来的设计）。

    显式写死而不是"抄上一行"——抄会把标题行的字号传染给正文行。
    """
    fixed = 0
    for offset in range(NOTES_ROWS):
        row = NOTES_FIRST + offset
        cell = sheet.cell(row, 1)
        if not cell.value:
            continue
        if offset == 0:
            target_font, target_align = NOTE_HEAD_FONT, Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            target_font, target_align = NOTE_BODY_FONT, NOTE_BODY_ALIGN
        if (cell.font.name, cell.font.size, cell.font.bold, cell.font.italic, cell.alignment.horizontal) \
                != (target_font.name, target_font.size, target_font.bold, target_font.italic, target_align.horizontal):
            cell.font = target_font
            cell.alignment = target_align
            fixed += 1
    return fixed


def refresh_notes(sheet) -> bool:
    """填写说明里点名的列必须与当前口径一致（删了跳出率/阅读次数/打开率就得跟着改）。"""
    changed = False
    for row in range(NOTES_FIRST, NOTES_FIRST + NOTES_ROWS):
        text = str(sheet[f"A{row}"].value or "")
        for prefix, replacement in NOTE_UPDATES:
            if text.startswith(prefix) and text != replacement:
                sheet[f"A{row}"] = replacement
                changed = True
    return changed


def inject_cached_results(path: str, values: dict) -> int:
    """把脚本算好的结果写进公式单元格的缓存值（`<f>` 后补 `<v>`）。

    openpyxl 写出的公式没有缓存值，WPS/Excel 在重算之前读不到数字——影响两件事：
    公式列一片空白、**条件格式的色阶着不上色**（色阶看的是单元格的值）。补上缓存值后，
    打开就能看到数字与色阶；Excel/WPS 仍会因为 fullCalcOnLoad 重算，公式照样是活的。

    values = {sheet 名: {"L6": 0.0072, ...}}，只处理数值型公式（文本型如匹配度跳过）。
    """
    sheets_needed = set(values)
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    sheet_files = sheet_xml_names(parts)
    patched = 0
    for sheet_name, cells in values.items():
        member = sheet_files.get(sheet_name)
        if not member or not cells:
            continue
        xml = parts[member].decode("utf-8")
        for address, number in cells.items():
            # openpyxl 写的是 <c ...><f>…</f><v></v></c>（空 v），要把它替换掉
            pattern = re.compile(r'(<c r="%s"[^>]*>)(<f>.*?</f>)(?:<v>[^<]*</v>)?(</c>)' % re.escape(address),
                                 re.S)
            def replace(match: re.Match, number=number) -> str:
                nonlocal patched
                patched += 1
                opening, formula, closing = match.group(1), match.group(2), match.group(3)
                if isinstance(number, str):  # 文本结果要标 t="str"
                    if ' t="' not in opening:
                        opening = f'{opening[:-1]} t="str">'
                    return f"{opening}{formula}<v>{escape(number)}</v>{closing}"
                return f"{opening}{formula}<v>{num_text(number)}</v>{closing}"
            xml = pattern.sub(replace, xml, count=1)
        parts[member] = xml.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return patched


def _attributes(tag: str) -> dict:
    """把 <xxx a="1" b="2"/> 拆成属性字典（不依赖属性顺序）。"""
    return dict(re.findall(r'([\w:]+)="([^"]*)"', tag))


def sheet_xml_names(parts: dict) -> dict:
    """从 workbook.xml + rels 解析出 {sheet 名: xl/worksheets/sheetN.xml}。

    注意别用"Id 在前 Target 在后"这种正则：openpyxl/Excel 写出的属性顺序不固定，
    Target 还可能是绝对路径（/xl/worksheets/sheet1.xml）。
    """
    targets = {}
    for tag in re.findall(r'<Relationship[^>]*>', parts["xl/_rels/workbook.xml.rels"].decode("utf-8")):
        attributes = _attributes(tag)
        if "Id" in attributes and "Target" in attributes:
            targets[attributes["Id"]] = attributes["Target"]
    names = {}
    for tag in re.findall(r'<sheet[^>]*>', parts["xl/workbook.xml"].decode("utf-8")):
        attributes = _attributes(tag)
        target = targets.get(attributes.get("r:id", ""), "")
        if attributes.get("name") and target:
            names[attributes["name"]] = "xl/" + target.lstrip("/").removeprefix("xl/")
    return names


def num_text(value: float) -> str:
    """缓存值写成十进制文本（避免科学计数法，Excel 对它更稳）。"""
    text = f"{float(value):.12f}".rstrip("0").rstrip(".")
    return text or "0"


def ensure_formulas(sheet) -> None:
    """整表统一写公式列（收藏率/分享率/匹配度/涨粉转化率），免得只有部分行是公式、口径打架。

    公式列不经过 merge_rows，格式要在这里一起设，否则涨粉转化率会显示成 0.0099 这种裸小数。
    """
    for row in range(FIRST_ROW, LAST_ROW + 1):
        for col, template in FORMULAS.items():
            cell = sheet[f"{col}{row}"]
            cell.value = template.replace("{r}", str(row))
            cell.number_format = PERCENT_FORMATS.get(col, "General")


def ensure_data_style(sheet, rows: range, columns: range, wrap: bool | None, height: float,
                      left_cols: tuple = ()) -> int:
    """数据区统一字体字号与对齐（数字数据一律居中），行高也统一。

    left_cols 里的列改为居左——成段文字（如透视表的结论）居中读起来费劲。
    只改外观，不动数值、边框、底纹；已经是目标样式的格子跳过，避免反复写样式表。
    """
    fixed = 0
    for row in rows:
        sheet.row_dimensions[row].height = height
        for column in columns:
            cell = sheet.cell(row, column)
            horizontal = "left" if get_column_letter(column) in left_cols else "center"
            font, align = cell.font, cell.alignment
            if (font.name, font.size, font.bold, align.horizontal, align.vertical, align.wrap_text) \
                    == (DATA_FONT, float(DATA_SIZE), False, horizontal, "center", wrap):
                continue
            cell.font = Font(name=DATA_FONT, size=DATA_SIZE)
            cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap)
            fixed += 1
    return fixed


def _scale_spec(sheet) -> list:
    """当前色阶规则的指纹，用来判断要不要重写。"""
    out = []
    for rng in sheet.conditional_formatting:
        rule = rng.rules[0]
        scale = rule.colorScale
        out.append((str(rng.sqref), rule.type, scale.cfvo[0].type, scale.cfvo[-1].type))
    return sorted(out)


def apply_color_scales(sheet) -> int:
    """给关键指标列加红-黄-绿色阶（低→高），让「哪些篇目高于/低于全表均值」一眼可见。

    透视表的诊断结论就是拿这几项跟全表均值比的，色阶是它的可视化对照。
    每次都按代码里的定义重建，保证改了阈值/颜色后能生效；内容一致时返回 0（不报变化）。
    """
    before = _scale_spec(sheet)
    sheet.conditional_formatting = ConditionalFormattingList()
    rules = ({col: COLOR_SCALE_RULE for col in COLOR_SCALE_COLS} | OUTLIER_RULES | REVERSE_RULES)
    for col, rule in rules.items():
        sheet.conditional_formatting.add(f"{col}{FIRST_ROW}:{col}{LAST_ROW}", ColorScaleRule(**rule))
    # 第六层逐行染色：每行 8 个渠道单独一条规则，行内比较（不是整列比较）
    first, last = COLUMNS[CHANNELS[0] + "占比"], COLUMNS[CHANNELS[-1] + "占比"]
    for row in range(FIRST_ROW, LAST_ROW + 1):
        sheet.conditional_formatting.add(f"{first}{row}:{last}{row}", ColorScaleRule(**CHANNEL_ROW_RULE))
    total = len(rules) + LAST_ROW - FIRST_ROW + 1
    return 0 if before == _scale_spec(sheet) else total


def extend_validations(sheet) -> None:
    """下拉菜单范围跟着数据区扩到 39 行。"""
    for validation in sheet.data_validations.dataValidation:
        ranges = list(validation.sqref.ranges)
        if not ranges:
            continue
        first = ranges[0]
        validation.sqref = (f"{get_column_letter(first.min_col)}{FIRST_ROW}:"
                            f"{get_column_letter(first.max_col)}{LAST_ROW}")


def ensure_capacity(sheet, needed: int) -> None:
    """数据区装不下就整体加长，填写说明继续往下让位。"""
    global LAST_ROW, NOTES_FIRST
    if needed <= LAST_ROW - FIRST_ROW + 1:
        return
    LAST_ROW = FIRST_ROW + needed + GROWTH_HEADROOM - 1
    NOTES_FIRST = LAST_ROW + 1
    migrate_layout(sheet)


def snapshot_rows(sheet) -> dict[str, dict[str, object]]:
    """按标题记下整行内容（含人工填的列），重排后原样还回去。"""
    rows: dict[str, dict[str, object]] = {}
    for row in range(FIRST_ROW, LAST_ROW + 1):
        title = sheet.cell(row, 3).value
        if not title:
            continue
        rows[str(title).strip()] = {col: sheet[f"{col}{row}"].value
                                   for col in DATA_COLUMNS if sheet[f"{col}{row}"].value is not None}
    return rows


def formula_values(values: dict) -> dict:
    """用 python 算出各公式的结果，供注入缓存值（数值 + 匹配度的文本）。"""
    out = {}
    for name, (numerator, denominator) in FORMULA_INPUTS.items():
        result = ratio(values.get(COLUMNS[numerator]), values.get(COLUMNS[denominator]))
        if result is not None:
            out[COLUMNS[name]] = result
    level = match_level(values.get(COLUMNS["女性读者占比"]), values.get(COLUMNS["25-45岁读者占比"]))
    if level:
        out[COLUMNS["目标人群匹配度"]] = level
    return out


def merge_rows(sheet, articles: list[dict]) -> dict:
    """增量并入：新标题追加行，已有标题只刷新数据列，历史行与人工列都不动。"""
    cached: dict[str, float] = {}
    incoming = {article["values"]["C"]: article["values"] for article in articles}
    history = snapshot_rows(sheet)
    dates = {title: values["B"] for title, values in incoming.items()}
    for title, values in history.items():
        dates.setdefault(title, str(values.get("B") or ""))
    titles = sorted(set(incoming) | set(history), key=lambda t: (dates[t], t), reverse=True)  # 日期倒序
    ensure_capacity(sheet, len(titles))

    for index, title in enumerate(titles):
        row = FIRST_ROW + index
        for column in range(1, LAST_COLUMN + 1):  # A~AB 整行清空（样式不动）
            sheet.cell(row, column).value = None
        sheet[f"A{row}"] = index + 1  # 序号按倒序重编
        # 本次有导出的行：人工/判断列从历史继承，其余列只认本次导出（旧布局/旧数值不会被带过来）；
        # raw/data/wechat/ 里已无导出的历史行：整行按原样保留
        if title in incoming:
            carried = {c: v for c, v in history.get(title, {}).items() if c in CARRY_COLS}
            values = {**carried, **incoming[title]}
        else:
            values = dict(history.get(title, {}))
        values["C"] = title
        for col, value in values.items():
            sheet[f"{col}{row}"] = stored(value)
            sheet[f"{col}{row}"].number_format = PERCENT_FORMATS.get(col, "General")
        for col, value in default_keys(title).items():  # 判断列：只填空位
            if not sheet[f"{col}{row}"].value:
                sheet[f"{col}{row}"] = value
        for col, value in formula_values(values).items():
            cached[f"{col}{row}"] = value
        match_col = COLUMNS["人群匹配度"]
        if not sheet[f"{match_col}{row}"].value:
            level = match_level(values.get(COLUMNS["女性读者占比"]), values.get(COLUMNS["25-45岁读者占比"]))
            if level:
                sheet[f"{match_col}{row}"] = level

    for row in range(FIRST_ROW + len(titles), LAST_ROW + 1):  # 空行清值，样式留着
        for column in range(1, LAST_COLUMN + 1):
            sheet.cell(row, column).value = None

    refreshed = {title: {col: (history[title].get(col), stored(value))
                         for col, value in incoming[title].items()
                         if not close(history[title].get(col), stored(value))}
                 for title in titles if title in history and title in incoming}
    return {
        "cached": cached,
        "titles": titles,
        "added": [title for title in titles if title not in history],
        "refreshed": {title: diff for title, diff in refreshed.items() if diff},
        "orphan": [title for title in titles if title not in incoming],  # raw/data/wechat/ 里已无导出，但行保留
    }


def collect_articles(sheet) -> list[dict]:
    """读数据区，取已导入文章的分组键与指标。比率一律由原始数字现算，不读单元格公式。"""
    articles = []
    for row in range(FIRST_ROW, LAST_ROW + 1):
        if not is_article_row(sheet, row):
            continue
        raw = {col: sheet[f"{col}{row}"].value for col in
               (COLUMNS["内容方向"], COLUMNS["核心价值类型"], COLUMNS["人群匹配度"], COLUMNS["阅读人数"],
                COLUMNS["公众号消息阅读人数"], COLUMNS["送达人数"], COLUMNS["推荐占比"], COLUMNS["完读率"],
                COLUMNS["分享人数"], COLUMNS["分享产生的阅读人数"],
                COLUMNS["女性读者占比"], COLUMNS["25-45岁读者占比"], COLUMNS["单篇涨粉数"],
                COLUMNS["目标人群阅读量"], COLUMNS["46岁以上占比"])}
        val = {col: v if isinstance(v, (int, float)) else None for col, v in raw.items()}
        articles.append({
            # 透视表列序为 内容方向 → 核心价值类型 → 匹配度，对应第一层 E → G → F
            "key": (raw[COLUMNS["内容方向"]] or UNFILLED, raw[COLUMNS["核心价值类型"]] or UNFILLED,
                    raw[COLUMNS["人群匹配度"]] or UNFILLED),
            # 订阅打开率在表里是公式（没有缓存值），这里用原始列现算
            "订阅打开率": ratio(val[COLUMNS["公众号消息阅读人数"]], val[COLUMNS["送达人数"]]),
            "完读率": val[COLUMNS["完读率"]],
            "分享率": ratio(val[COLUMNS["分享人数"]], val[COLUMNS["阅读人数"]]),
            "女性占比": val[COLUMNS["女性读者占比"]],
            "25-45岁占比": val[COLUMNS["25-45岁读者占比"]],
            "涨粉转化率": ratio(val[COLUMNS["单篇涨粉数"]], val[COLUMNS["阅读人数"]]),
            "推荐占比": val[COLUMNS["推荐占比"]],
            "阅读人数": val[COLUMNS["阅读人数"]],
            # 扩散系数在表里是公式（无缓存值），用原始列现算，和分享率一个路子
            "分享扩散系数": ratio(val[COLUMNS["分享产生的阅读人数"]], val[COLUMNS["分享人数"]]),
            "女性读者占比": val[COLUMNS["女性读者占比"]],
            "25-45岁读者占比": val[COLUMNS["25-45岁读者占比"]],
            "46岁以上占比": val[COLUMNS["46岁以上占比"]],
            "目标人群阅读量": val[COLUMNS["目标人群阅读量"]],
            "单篇涨粉数": val[COLUMNS["单篇涨粉数"]],
            "发布日期": sheet[f"B{row}"].value,
        })
    return articles


def precise_fans(article: dict) -> float:
    """精准涨粉 = 涨粉 × 25-45岁占比，即这次涨粉里落在目标年龄段的那部分。"""
    return (article.get("单篇涨粉数") or 0) * (article.get("25-45岁读者占比") or 0)


# ---------- 结论框架 v2：先分流量层，再在同层内诊断 ----------
# 为什么必须分层：推荐流量是冷流量，完读率/分享率天然比私域低（实测算法型完读中位 20.7%、私域型 46.9%）。
# 用同一条线比较两类文章，结论必然反向——旧版就把「阅读 38596 的算法爆文」判成了「减少或优化」。
ALGO_RECOMMEND, PRIVATE_RECOMMEND = 0.7, 0.3
REGIMES = ("算法型", "混合型", "私域型")
DIMENSIONS = (("内容力", "完读率"), ("传播力", "分享率"), ("扩散力", "分享扩散系数"), ("沉淀力", "涨粉转化率"))
WEAK_LABEL = {"内容力": "内容承接偏弱", "传播力": "传播偏弱", "扩散力": "扩散偏弱", "沉淀力": "沉淀偏弱"}
BASELINE_METRICS = ("订阅打开率", "完读率", "分享率", "分享扩散系数", "涨粉转化率", "女性读者占比", "25-45岁读者占比")
BASELINE_KEYS = BASELINE_METRICS  # 与 collect_articles 的键名一致，避免"列名 vs 字典键"混淆


def regime_of(article: dict) -> str:
    """按推荐占比分流量层：≥70% 算法型、≤30% 私域型、中间混合型。"""
    recommend = article.get("推荐占比")
    if recommend is None:
        return "混合型"
    if recommend >= ALGO_RECOMMEND:
        return "算法型"
    return "私域型" if recommend <= PRIVATE_RECOMMEND else "混合型"


def regime_baselines(articles: list[dict]) -> dict:
    """每一层的基准。比率用中位数（抗离群值），涨粉率中位常为 0 所以用均值。"""
    out = {}
    for regime in REGIMES:
        items = [a for a in articles if regime_of(a) == regime]
        stats = {name: median([a[name] for a in items if a.get(name) is not None])
                 for name in BASELINE_METRICS}
        stats["涨粉转化率"] = mean([a["涨粉转化率"] for a in items if a.get("涨粉转化率") is not None])
        stats["阅读人数"] = median([a["阅读人数"] for a in items if a.get("阅读人数") is not None])
        stats["文章数"] = len(items)
        stats["推荐占比"] = median([a["推荐占比"] for a in items if a.get("推荐占比") is not None])
        out[regime] = stats
    return out


ACTIONS = {
    ("增长轨", "内容力"): "算法流量全靠开头承接——把前 3 行改成场景或冲突切入，别先讲道理",
    ("增长轨", "传播力"): "冷流量读完就走——结尾放一句可转发的主张，把算法来的读者变成分享者",
    ("增长轨", "扩散力"): "转出去带不动人——文末给一个具体的转发理由（清单或结论卡片）",
    ("增长轨", "沉淀力"): "选题偏窄或引导弱——换更普适的钩子，文末补关注理由",
    ("增长轨", ""): "算法已经接住——保持选题方向，把完读率与关注转化当下一阶段指标",
    ("基本盘轨", "内容力"): "私域读者愿意点开却读不完——中段节奏问题：缩短段落、前置结论",
    ("基本盘轨", "传播力"): "读完不转发——加一句可转发的主张，或把结论做成清单",
    ("基本盘轨", "扩散力"): "转发动机弱或带不动人——给分享者一个能带走的金句或清单",
    ("基本盘轨", "沉淀力"): "留不住人——文末补关注理由，并把该选题做成 2~3 篇系列",
    ("基本盘轨", ""): "私域基本盘稳——保持节奏，试探能不能被算法接住（改标题、换首发时段）",
}
AUDIENCE_ACTION = "内容与传播都不弱，卡在人群匹配——切角换到 25-45 岁女性的具体处境再测"
NO_PROFILE_ACTION = "后台尚未生成用户画像（阅读量太小）——先攒阅读量，再判人群匹配"


def conclusion(items: list[dict], baselines: dict) -> str:
    """每组的结论（v2）：流量结构定位 ＋ 同层基准诊断 ＋ 按轨道给动作。

    判定不再用一套绝对达标线，而是：① 先看这一组主要是哪种流量结构；② 每个指标跟**同层**文章的
    典型水平比（同层期望＝组内各篇所属层基准的均值）；③ 双轨给动作——增长轨（算法型）盯冷流量的
    承接与关注转化，基本盘轨（私域）盯打开、分享与扩散。
    """
    count = len(items)
    stats = {name: mean([a[name] for a in items if a.get(name) is not None])
             for name in METRICS + EXTRA_METRICS}
    regimes = Counter(regime_of(a) for a in items)
    regime = regimes.most_common(1)[0][0]
    track = "增长轨" if regime == "算法型" else "基本盘轨"
    expected = {name: mean([baselines[regime_of(a)].get(name) for a in items
                            if baselines[regime_of(a)].get(name) is not None])
                for name in BASELINE_METRICS}

    gaps = {label: stats.get(metric, 0) / expected[metric]
            for label, metric in DIMENSIONS
            if stats.get(metric) is not None and expected.get(metric)}
    weakest = min(gaps, key=gaps.get) if gaps else None
    if weakest and gaps[weakest] >= 0.90:  # 跟同层差不多就别硬找短板（0.90 倍以上视作持平）
        weakest = None

    female, young = stats.get("女性占比"), stats.get("25-45岁占比")
    if female is None or young is None:
        action = NO_PROFILE_ACTION
    elif not weakest and ((female / expected["女性占比"] < 0.9 if expected.get("女性占比") else False)
                          or (young / expected["25-45岁占比"] < 0.9 if expected.get("25-45岁占比") else False)):
        action = AUDIENCE_ACTION  # 各维度不低于同层却人群偏：别再改内容，改切角
    else:
        action = ACTIONS[(track, weakest or "")]

    structure = f"{regime}为主 {regimes[regime]}/{count} 篇·{track}"
    if stats.get("推荐占比") is not None:
        structure += f"（推荐占比 {pct(stats['推荐占比'])}）"
    head = f"样本{count}篇｜" if count < 3 else ""
    weakest_text = WEAK_LABEL[weakest] + f"（只有同层的 {gaps[weakest]:.0%}）" if weakest else "各维度不低于同层"
    items_to_show = [("完读率", "完读率"), ("分享率", "分享率"), ("涨粉转化率", "涨粉转化率")]
    if weakest == "扩散力":
        items_to_show.insert(2, ("扩散系数", "分享扩散系数"))
    diagnosis = "、".join(
        (f"{title} {stats.get(metric):.2f}" if metric == "分享扩散系数" else f"{title} {pct(stats.get(metric))}")
        + (f"（同层 {expected[metric]:.2f}）" if metric == "分享扩散系数" and expected.get(metric)
           else f"（同层 {pct(expected[metric])}）" if expected.get(metric) else "")
        for title, metric in items_to_show if stats.get(metric) is not None)
    return (f"【定位】{head}{structure}｜{weakest_text}\n"
            f"【诊断】{diagnosis}\n"
            f"【动作】{action}")


REGIME_ROLE = {"算法型": "流量入口", "混合型": "过渡层", "私域型": "信任底盘"}
REGIME_ACTIONS = {
    "算法型": "把冷流量的承接做厚：前 3 行场景化、文末给关注理由；选题继续找面足够宽的钩子（爆款的共性是选题普适）",
    "混合型": "推荐与私域各半——先把私域的打开率与扩散做起来（标题、可转发主张），再放大算法种子",
    "私域型": "私域是信任底盘：打开率与扩散系数是核心指标；这类阅读小但分享率高，别用阅读量否定它",
}


def regime_conclusion(regime: str, stats: dict, layer_reads: int, total_reads: int, table_mean_grow: float | None) -> str:
    """层内结论：这层在全表里的角色 + 诊断 + 该层该怎么用。"""
    contribution = layer_reads / total_reads if total_reads else 0
    grow = stats.get("涨粉转化率")
    grow_text = ""
    if grow is not None and table_mean_grow:
        grow_text = f"；涨粉 {pct(grow)}（全表均值 {pct(table_mean_grow)}，{'更高' if grow > table_mean_grow else '更低'}）"
    return (f"【定位】{regime}·{REGIME_ROLE[regime]}｜{stats['文章数']} 篇，占全表阅读量 {contribution:.0%}\n"
            f"【诊断】中位阅读 {stats.get('阅读人数') or 0:.0f}、完读 {pct(stats.get('完读率'))}"
            f"、分享 {pct(stats.get('分享率'))}、扩散 {stats.get('分享扩散系数') or 0:.2f}{grow_text}\n"
            f"【动作】{REGIME_ACTIONS[regime]}")


def build_regime_sheet(workbook, articles: list[dict], baselines: dict) -> list[tuple]:
    """新建/更新「流量结构透视」：三层各自的中位表现 + 层内结论。

    这张表回答的是「哪种流量结构值得投入」——它跟「方向×价值类型」那张回答的不是同一个问题，
    所以单开一张，标题行样式沿用透视表那张（改配色/字体只需改一处）。
    """
    pivot = workbook[PIVOT_SHEET]
    if REGIME_SHEET in workbook.sheetnames:
        sheet = workbook[REGIME_SHEET]
    else:
        sheet = workbook.create_sheet(REGIME_SHEET, workbook.sheetnames.index(PIVOT_SHEET) + 1)
    headers = ("流量结构", "文章数", "中位阅读人数", "中位订阅打开率", "中位完读率", "中位分享率",
               "中位分享扩散系数", "中位女性占比", "中位25-45岁占比", "中位涨粉转化率", "结论与建议")
    style_from = ("A", "D", "D", "E", "F", "G", "G", "H", "I", "J", "K")  # 逐列借用透视表的表头样式
    for column, (title, source) in enumerate(zip(headers, style_from), start=1):
        cell = sheet.cell(4, column)
        cell.value = title
        cell._style = copy(pivot[f"{source}4"]._style)
    sheet.cell(1, 1).value = "流量结构 × 数据表现（分层基准）"
    sheet.cell(1, 1)._style = copy(pivot["A1"]._style)
    sheet.cell(2, 1).value = ("按推荐占比分层：算法型 ≥70% / 混合型 30~70% / 私域型 ≤30%。"
                              "同层才有可比性——算法流量是冷流量，完读率、分享率天然低于私域，"
                              "所以这里取层内中位数，只用来判断「比同类文章好还是差」")
    sheet.cell(2, 1)._style = copy(pivot["A2"]._style)

    total_reads = sum(a.get("阅读人数") or 0 for a in articles)
    table_mean_grow = mean([a["涨粉转化率"] for a in articles if a.get("涨粉转化率") is not None])
    layer_reads = {regime: sum(a.get("阅读人数") or 0 for a in articles if regime_of(a) == regime)
                   for regime in REGIMES}
    rows = []
    for index, regime in enumerate(REGIMES):
        stats = baselines[regime]
        row = 5 + index
        values = (f"{regime}（推荐{'≥70%' if regime == '算法型' else '≤30%' if regime == '私域型' else '30~70%'}）",
                  stats["文章数"], stats.get("阅读人数"), stats.get("订阅打开率"), stats.get("完读率"),
                  stats.get("分享率"), stats.get("分享扩散系数"), stats.get("女性读者占比"),
                  stats.get("25-45岁读者占比"), stats.get("涨粉转化率"))
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row, column)
            cell.value = value
            cell._style = copy(pivot.cell(5, column)._style)
            if column in (4, 5, 6, 9):
                cell.number_format = "0.0%"
            elif column == 7:
                cell.number_format = "0.00"
            elif column == 8:
                cell.number_format = "0.0%"
            elif column == 10:
                cell.number_format = "0.00%"
        text = regime_conclusion(regime, stats, layer_reads[regime], total_reads, table_mean_grow)
        cell = sheet.cell(row, 11)
        cell.value = text
        cell._style = copy(pivot.cell(5, 11)._style)
        rows.append((regime, stats, text))

    note_row = 5 + len(REGIMES) + 1
    sheet.cell(note_row, 1).value = ("说明：中位数代表「典型文章」，比均值抗离群值（全表阅读均值 1436、中位仅 116）；"
                                     "涨粉转化率用层内均值（中位常为 0）。分层是为了解决「推荐流量越大，"
                                     "完读率与分享率必然越低」导致的误判——旧版用一条绝对线比较两类文章，"
                                     "把阅读 38596 的算法爆文判成了「减少或优化」")
    sheet.cell(note_row, 1)._style = copy(pivot.cell(29, 1)._style)
    for merged in (f"A2:{get_column_letter(len(headers))}2",
                   f"A{note_row}:{get_column_letter(len(headers))}{note_row}"):
        sheet.merge_cells(merged)
    for column, width in zip("ABCDEFGHIJK", (26, 8, 12, 13, 10, 10, 13, 12, 14, 12, CONCLUSION_WIDTH)):
        sheet.column_dimensions[column].width = width
    for row, height in ((1, 35), (2, 30), (4, 30), (5, 76), (6, 76), (7, 76), (note_row, 34)):
        sheet.row_dimensions[row].height = height
    return rows


def rebuild_pivot(workbook) -> list[tuple]:
    """按 内容方向×核心价值类型×人群匹配度 聚合「文章数据记录」，重写透视表数据区。"""
    articles = collect_articles(workbook[SHEET])
    baselines = regime_baselines(articles)  # 分层基准：同一层内部才能互相比
    grouped: dict[tuple, list[dict]] = {}
    for article in articles:
        grouped.setdefault(article["key"], []).append(article)

    rows = []
    for key, items in sorted(grouped.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        stats = {name: mean([i[name] for i in items if i.get(name) is not None])
                 for name in METRICS + EXTRA_METRICS}
        # K/L 两列：这一组实际触达了多少目标读者、换回多少精准粉（篇数少时比均值更说明问题）
        reach = sum(i.get("目标人群阅读量") or 0 for i in items)
        precise = sum(precise_fans(i) for i in items)
        rows.append((key, len(items), stats, conclusion(items, baselines), reach, precise))
    capacity = PIVOT_LAST_ROW - PIVOT_FIRST_ROW + 1
    if len(rows) > capacity:
        raise SystemExit(f"透视表数据区只有 {capacity} 行，分组数 {len(rows)} 超出")

    pivot = workbook[PIVOT_SHEET]
    for index, (key, count, stats, text, reach, precise) in enumerate(rows):
        row = PIVOT_FIRST_ROW + index
        for col, value in zip("ABC", key):
            pivot[f"{col}{row}"] = value
        pivot[f"D{row}"] = count
        for col, name in zip("EFGHIJ", METRICS):
            cell = pivot[f"{col}{row}"]
            cell.value = stats[name]
            cell.number_format = "0.0%"
        pivot[f"K{row}"] = reach
        pivot[f"K{row}"].number_format = "#,##0"
        pivot[f"L{row}"] = round(precise, 1)
        pivot[f"L{row}"].number_format = "0.0"
        pivot[f"M{row}"] = text

    for row in range(PIVOT_FIRST_ROW + len(rows), PIVOT_LAST_ROW + 1):
        for col in "ABCDEFGHIJKLM":
            pivot[f"{col}{row}"] = None
    return rows


def ensure_pivot_columns(sheet) -> bool:
    """透视表在结论列前新增两列：目标人群阅读合计（K）、精准涨粉（L），结论列由 K 移到 M。

    表头与合并区都要跟着扩，否则新列落在合并区外会显示错乱。写成幂等的：重复跑只改一次。
    """
    changed = False
    if sheet["M4"].value != "结论与建议":  # 先搬结论列，此时 K4 还是老表头，样式从它借
        sheet["M4"].value = "结论与建议"
        sheet["M4"]._style = copy(sheet["K4"]._style)
        changed = True
    for col, title in (("K", "目标人群阅读合计"), ("L", "精准涨粉")):
        if sheet[f"{col}4"].value != title:
            sheet[f"{col}4"].value = title
            sheet[f"{col}4"]._style = copy(sheet["J4"]._style)
            changed = True
    existing = {str(item) for item in sheet.merged_cells.ranges}
    for row in (1, 2, PIVOT_LAST_ROW + 1, PIVOT_LAST_ROW + 2, PIVOT_LAST_ROW + 3, PIVOT_LAST_ROW + 4):
        old, new = f"A{row}:K{row}", f"A{row}:M{row}"
        if old in existing:
            sheet.unmerge_cells(old)
            sheet.merge_cells(new)
            changed = True
    for col, width in (("K", 16), ("L", 10), ("M", CONCLUSION_WIDTH)):
        sheet.column_dimensions[col].width = width
    return changed


def build_month_sheet(workbook, articles: list[dict]) -> list[tuple]:
    """新建/更新「月度趋势」：按发布月汇总，看三个目标是否在靠近。

    时间正序（旧的在上）——这是唯一一张看走向的表，倒序看不出趋势。
    """
    pivot = workbook[PIVOT_SHEET]
    if MONTH_SHEET in workbook.sheetnames:
        sheet = workbook[MONTH_SHEET]
    else:
        sheet = workbook.create_sheet(MONTH_SHEET, workbook.sheetnames.index(REGIME_SHEET) + 1)
    headers = ("月份", "篇数", "阅读合计", "目标人群阅读合计", "中位订阅打开率",
               "中位25-45岁占比", "中位46岁以上占比", "涨粉合计", "精准涨粉")
    style_from = ("A", "D", "C", "C", "D", "I", "I", "C", "C")
    for column, (title, source) in enumerate(zip(headers, style_from), start=1):
        cell = sheet.cell(4, column)
        cell.value = title
        cell._style = copy(pivot[f"{source}4"]._style)
    sheet.cell(1, 1).value = "月度趋势（按发布月，时间正序）"
    sheet.cell(1, 1)._style = copy(pivot["A1"]._style)
    sheet.cell(2, 1).value = ("三个目标看是否在靠近：中位订阅打开率（≥5%）、中位46岁以上占比（反向护栏 ≤40%）、"
                              "精准涨粉＝涨粉×25-45岁占比。用中位数而非均值——单篇爆文会把均值拉飞，看不出走向")
    sheet.cell(2, 1)._style = copy(pivot["A2"]._style)

    by_month: dict[str, list[dict]] = {}
    for article in articles:
        month = str(article.get("发布日期") or "")[:7]
        if len(month) == 7:
            by_month.setdefault(month, []).append(article)

    rows = []
    for month in sorted(by_month):
        items = by_month[month]
        reads = sum(i.get("阅读人数") or 0 for i in items)
        reach = sum(i.get("目标人群阅读量") or 0 for i in items)
        fans = sum(i.get("单篇涨粉数") or 0 for i in items)
        precise = sum(precise_fans(i) for i in items)
        rows.append((month, len(items), reads, reach,
                     median([i["订阅打开率"] for i in items if i.get("订阅打开率") is not None]),
                     median([i["25-45岁占比"] for i in items if i.get("25-45岁占比") is not None]),
                     median([i["46岁以上占比"] for i in items if i.get("46岁以上占比") is not None]),
                     fans, round(precise, 1)))
    for index, values in enumerate(rows):
        row = 5 + index
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row, column)
            cell.value = value
            cell._style = copy(pivot.cell(5, (1, 4, 3, 3, 4, 9, 9, 3, 3)[column - 1])._style)
            if column == 5:
                cell.number_format = "0.0%"
            elif column in (6, 7):
                cell.number_format = "0.0%"
            elif column in (3, 4):
                cell.number_format = "#,##0"
            elif column == 9:
                cell.number_format = "0.0"
    last = 5 + len(rows)
    # 先拆掉上次留下的说明合并区：合并后除左上格外都是只读的 MergedCell，直接清空会报错
    for merged in [str(item) for item in sheet.merged_cells.ranges]:
        if not merged.startswith("A2:"):
            sheet.unmerge_cells(merged)
    for row in range(last, last + 8):  # 清掉上次跑留下的月份行与旧说明
        for column in range(1, len(headers) + 1):
            sheet.cell(row, column).value = None
    note_row = last + 1
    sheet.cell(note_row, 1).value = ("说明：月份按发布日期归月；篇数=1 的月份中位数即该篇自身，别过度解读。"
                                     "若某月只有算法型文章，46岁以上占比会明显偏高——那是流量结构造成的，"
                                     "不是账号在变老（详见「流量结构透视」）")
    sheet.cell(note_row, 1)._style = copy(pivot.cell(29, 1)._style)
    for merged in (f"A2:{get_column_letter(len(headers))}2",
                   f"A{note_row}:{get_column_letter(len(headers))}{note_row}"):
        if merged not in {str(item) for item in sheet.merged_cells.ranges}:
            sheet.merge_cells(merged)
    for column, width in zip("ABCDEFGHI", (14, 8, 12, 16, 15, 16, 16, 11, 10)):
        sheet.column_dimensions[column].width = width
    sheet.row_dimensions[1].height = 35
    sheet.row_dimensions[2].height = 30
    sheet.row_dimensions[4].height = 30
    for index in range(len(rows)):
        sheet.row_dimensions[5 + index].height = DATA_ROW_HEIGHT
    sheet.row_dimensions[note_row].height = 34
    return rows


def dedupe(articles: list[dict]) -> list[dict]:
    """同一篇导出多次（文件名带 (1) 那种）：保留指标更全的一份。"""
    best: dict[str, dict] = {}
    for article in articles:
        title = article["values"]["C"]
        filled = sum(v is not None for v in article["values"].values())
        if title not in best or filled > best[title][0]:
            best[title] = (filled, article)
    return [article for _, article in best.values()]


def check_not_open() -> None:
    """表正被 WPS/Excel 打开时拒绝写入。

    办公软件打开着这个文件时，它内存里拿的是旧副本；脚本写完后你在软件里一保存，
    旧内容就会把脚本写的结果整片覆盖（真发生过一次：列结构、结论、样式全被回退成旧版）。
    """
    if not os.path.exists(OUTPUT):
        return
    lock = os.path.join(os.path.dirname(OUTPUT), ".~" + os.path.basename(OUTPUT))
    if os.path.exists(lock):
        raise SystemExit(
            f"检测到锁文件 {os.path.basename(lock)}：{os.path.basename(OUTPUT)} 正被 WPS/Excel 打开。\n"
            "请先关闭它（不要保存旧内容）再运行脚本，否则你一保存就会覆盖本次写入。\n"
            "确认其实没打开（上次崩溃留下的锁文件）时，删掉这个 .~ 开头的文件再跑。")


def main() -> None:
    check_not_open()
    sources = sorted(glob.glob(os.path.join(DATA_DIR, "*.xls")))
    if not sources:
        raise SystemExit(f"{DATA_DIR} 下没有 .xls")

    articles, failures = [], []
    for path in sources:
        try:
            articles.append(read_export(path))
        except BaseException as error:  # 单份文件坏掉不该拖垮整批
            failures.append((os.path.basename(path), str(error)))
    articles = dedupe(articles)
    if not articles or failures:  # 一份解析不了就整体停下，别把表改成一个半成品
        for name, error in failures:
            print(f"解析失败: {name}\n  {error}")
        raise SystemExit(f"成功解析 {len(articles)} 篇、失败 {len(failures)} 份，未改动输出文件")

    fresh = not os.path.exists(OUTPUT)
    if fresh:
        shutil.copy2(TEMPLATE, OUTPUT)  # 模板本体不动
    workbook = openpyxl.load_workbook(OUTPUT)
    sheet = workbook[SHEET]
    if fresh:  # 模板示例行的判断列/人工列先清空，免得示例值被当成真实输入
        for row in SAMPLE_ROWS:
            for col in SAMPLE_CLEAR_COLS:
                sheet[f"{col}{row}"] = None

    moved = migrate_layout(sheet)
    layout_written = write_layout(sheet)
    notes_refreshed = ensure_channel_note(sheet) or refresh_notes(sheet)
    documented = document_basis(workbook[PIVOT_SHEET])
    columns_added = ensure_pivot_columns(workbook[PIVOT_SHEET])
    result = merge_rows(sheet, articles)
    ensure_formulas(sheet)
    restyled = ensure_data_style(sheet, range(FIRST_ROW, LAST_ROW + 1), range(1, LAST_COLUMN + 1),
                                 wrap=None, height=DATA_ROW_HEIGHT)
    scales = apply_color_scales(sheet)
    extend_validations(sheet)
    pivot_rows = rebuild_pivot(workbook)
    regime_rows = build_regime_sheet(workbook, collect_articles(sheet), regime_baselines(collect_articles(sheet)))
    month_rows = build_month_sheet(workbook, collect_articles(sheet))
    restyled += ensure_data_style(workbook[PIVOT_SHEET], range(PIVOT_FIRST_ROW, PIVOT_LAST_ROW + 1),
                                  range(1, 14), wrap=PIVOT_WRAP, height=PIVOT_ROW_HEIGHT,
                                  left_cols=PIVOT_LEFT_COLS)
    workbook.calculation.fullCalcOnLoad = True  # 让 WPS/Excel 打开时重算公式列
    workbook.save(OUTPUT)
    injected = inject_cached_results(OUTPUT, {SHEET: result["cached"]})

    print(f"raw/data/wechat/ 下 {len(sources)} 份导出 → 新文章 {len(result['added'])} 篇、"
          f"刷新数据 {len(result['refreshed'])} 篇、历史保留 {len(result['orphan'])} 篇"
          f"（表内共 {len(result['titles'])} 行，按发布日期倒序）")
    if moved:
        print(f"表结构已迁移：数据区 {FIRST_ROW}~{LAST_ROW} 行，"
              f"填写说明下移到 {NOTES_FIRST}~{NOTES_FIRST + NOTES_ROWS - 1} 行")
    if layout_written:
        print("表头已按新口径重铺：" + "、".join(
            [f"{banner.split('：')[1]}({len(names)}列)" for banner, names in LAYERS]))
    if documented:
        print("透视表说明文字已更新为 v2 口径（同层基准）")
    if columns_added:
        print("透视表已加两列：目标人群阅读合计(K) / 精准涨粉(L)，结论列移到 M")
    if notes_refreshed:
        print("填写说明里的列名已同步（去掉跳出率/阅读次数/打开率）")
    if restyled:
        print(f"外观已统一：{restyled} 个数据格改为 {DATA_FONT} {DATA_SIZE} 居中（原来是默认宋体/未居中）")
    if injected:
        print(f"已给 {injected} 个公式单元格补上缓存值（否则 WPS 重算前读不到数，色阶着不上色）")
    if scales:
        cols = "/".join(list(COLOR_SCALE_COLS) + list(OUTLIER_RULES) + list(REVERSE_RULES))
        print(f"已给关键指标列加色阶：{cols}（红→黄→绿，低→高；阅读人数按分位数，免受离群值压色）"
              f"；第六层另按行染色（每行 8 个渠道行内比较，浅→深蓝）")
    for title in result["added"]:
        print(f"  ＋新增: {title[:40]}")
    for title, diff in result["refreshed"].items():
        changes = " / ".join(f"{col}: {old} → {new}" for col, (old, new) in list(diff.items())[:4])
        print(f"  ↻刷新: {title[:26]} → {changes}")
    for title in result["orphan"]:
        print(f"  ＝保留（raw/data/wechat/ 里已无这份导出）: {title[:40]}")

    print(f"\n{'序号':<4}{'发布日期':<12}{'阅读':>6}{'订阅打开率':>10}{'完读率':>8}{'分享率':>8}"
          f"  内容方向 / 价值类型 / 匹配度  标题")
    for index, title in enumerate(result["titles"]):
        row = FIRST_ROW + index

        def cell(name: str):
            return sheet[f'{COLUMNS[name]}{row}'].value

        reads = cell("阅读人数")
        print(f"{index + 1:<5}{str(sheet[f'B{row}'].value):<12}{str(reads if reads is not None else '—'):>6}"
              f"{pct(ratio(cell('公众号消息阅读人数'), cell('送达人数'))):>10}"
              f"{pct(cell('完读率')):>8}{pct(ratio(cell('分享人数'), reads)):>8}"
              f"  {cell('内容方向')} / {cell('核心价值类型')} / {cell('人群匹配度') or UNFILLED}  {title[:26]}")

    print(f"\n「{REGIME_SHEET}」{len(regime_rows)} 层：")
    for regime, layer_stats, _ in regime_rows:
        print(f"  {regime}｜{layer_stats['文章数']} 篇｜中位阅读 {layer_stats.get('阅读人数') or 0:.0f}"
              f"｜完读 {pct(layer_stats.get('完读率'))}｜分享 {pct(layer_stats.get('分享率'))}"
              f"｜扩散 {layer_stats.get('分享扩散系数') or 0:.2f}｜涨粉 {pct(layer_stats.get('涨粉转化率'))}")

    print(f"\n「{MONTH_SHEET}」{len(month_rows)} 个月（时间正序）：")
    for month, count, reads, reach, open_rate, young, old, fans, precise in month_rows:
        print(f"  {month}｜{count} 篇｜阅读 {reads:>6,}｜目标阅读 {reach:>7,.0f}"
              f"｜中位打开 {pct(open_rate):>6}｜25-45 中位 {pct(young):>6}｜46+ 中位 {pct(old):>6}"
              f"｜涨粉 {fans:>3.0f}｜精准涨粉 {precise:>4.1f}")

    print(f"\n「{PIVOT_SHEET}」{len(pivot_rows)} 个分组：")
    for key, count, stats, text, reach, precise in pivot_rows:
        print(f"  {' × '.join(key)} | {count} 篇 | 订阅打开率 {pct(stats['订阅打开率'])}"
              f" | 完读率 {pct(stats['完读率'])} | 分享率 {pct(stats['分享率'])}"
              f" | 女性 {pct(stats['女性占比'])} | 25-45岁 {pct(stats['25-45岁占比'])}"
              f" | 涨粉转化 {pct(stats['涨粉转化率'])} | 目标阅读 {reach:>6,.0f} | 精准涨粉 {precise:>4.1f}"
              f" | {text.splitlines()[0]}")
    for article in articles:
        for warning in article["warnings"]:
            print(f"告警: {article['file']} → {warning}")
    print(f"\n输出: {os.path.basename(OUTPUT)}")
    print("脚本从不写: " + " / ".join(NEVER_WRITTEN)
          + " | 空着才填默认值: " + " / ".join(JUDGMENT_COLS))


if __name__ == "__main__":
    main()
