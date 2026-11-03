# -*- coding: utf-8 -*-
"""生成六级 6 周背单词计划：words.json + 打卡表 xlsx"""
import json, os
from datetime import date, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)  # CET6-Plan 主目录
START = date(2026, 11, 2)
EXAM = date(2026, 12, 12)
INTERVALS = [1, 4, 15]
PER_DAY = 10

# (单词, 词性, 中文释义, 常用搭配/短句)
W = [
("advocate", "v.", "提倡，主张", "advocate reform 主张改革"),
("assess", "v.", "评估，评定", "assess the risk 评估风险"),
("assume", "v.", "假定；承担", "assume responsibility 承担责任"),
("attribute", "v.", "把……归因于", "attribute success to luck 把成功归因于运气"),
("acknowledge", "v.", "承认；致谢", "acknowledge the fact 承认事实"),
("alter", "v.", "改变，更改", "alter the plan 改变计划"),
("accelerate", "v.", "加速，促进", "accelerate growth 加速增长"),
("diminish", "v.", "减少，削弱", "diminish in value 贬值"),
("enhance", "v.", "提高，增强", "enhance efficiency 提高效率"),
("deteriorate", "v.", "恶化，变坏", "health deteriorated 健康恶化了"),
("articulate", "v.", "清晰地表达", "articulate one's views 表达观点"),
("convey", "v.", "传达，表达；运送", "convey a message 传达信息"),
("persuade", "v.", "说服，劝说", "persuade sb. into doing 说服某人做"),
("contradict", "v.", "反驳；与……矛盾", "contradict oneself 自相矛盾"),
("elaborate", "v.", "详细阐述", "elaborate on the point 详述这一点"),
("impose", "v.", "强加；征收", "impose a ban on 对……实施禁令"),
("restrain", "v.", "抑制，约束", "restrain oneself 克制自己"),
("violate", "v.", "违反，侵犯", "violate the law 违法"),
("comply", "v.", "遵守，服从", "comply with rules 遵守规则"),
("enforce", "v.", "执行，强制实施", "enforce the law 执法"),
("allocate", "v.", "分配，拨给", "allocate funds 分配资金"),
("compensate", "v.", "补偿，赔偿", "compensate for the loss 赔偿损失"),
("consume", "v.", "消耗；消费", "consume energy 消耗能量"),
("distribute", "v.", "分发，分配", "distribute resources 分配资源"),
("invest", "v.", "投资；投入", "invest in education 投资教育"),
("conceive", "v.", "构想，设想；怀孕", "conceive of an idea 构想一个主意"),
("deduce", "v.", "推断，演绎", "deduce from evidence 从证据推断"),
("distinguish", "v.", "区分，辨别", "distinguish A from B 区分A和B"),
("perceive", "v.", "察觉，认为", "perceive as a threat 视为威胁"),
("speculate", "v.", "推测；投机", "speculate on the outcome 推测结果"),
("accomplish", "v.", "完成，实现", "accomplish a goal 实现目标"),
("endeavor", "v./n.", "努力，尽力", "endeavor to improve 努力改进"),
("undertake", "v.", "承担，着手", "undertake a project 承担项目"),
("pursue", "v.", "追求，从事", "pursue a career 追求事业"),
("attain", "v.", "达到，获得", "attain success 获得成功"),
("confront", "v.", "面对，正视", "confront a challenge 直面挑战"),
("cope", "v.", "应对，处理", "cope with stress 应对压力"),
("tackle", "v.", "处理，解决", "tackle the problem 解决问题"),
("resolve", "v.", "解决；决心", "resolve a conflict 解决冲突"),
("endure", "v.", "忍受；持续", "endure hardship 忍受艰辛"),
("consequence", "n.", "后果，结果", "as a consequence of 由于"),
("circumstance", "n.", "情况，环境", "under no circumstances 决不"),
("phenomenon", "n.", "现象", "a social phenomenon 社会现象"),
("tendency", "n.", "趋势，倾向", "a tendency to do 做……的倾向"),
("discrepancy", "n.", "差异，不一致", "a discrepancy between 两者间的差异"),
("consensus", "n.", "共识，一致意见", "reach a consensus 达成共识"),
("controversy", "n.", "争议，争论", "a heated controversy 激烈争议"),
("dilemma", "n.", "困境，两难", "face a dilemma 面临两难"),
("prejudice", "n.", "偏见，成见", "racial prejudice 种族偏见"),
("discrimination", "n.", "歧视；辨别力", "gender discrimination 性别歧视"),
("budget", "n.", "预算", "a tight budget 紧张的预算"),
("revenue", "n.", "收入，税收", "tax revenue 税收"),
("expenditure", "n.", "支出，开销", "public expenditure 公共支出"),
("inflation", "n.", "通货膨胀", "curb inflation 抑制通胀"),
("prosperity", "n.", "繁荣，兴旺", "economic prosperity 经济繁荣"),
("competence", "n.", "能力，胜任", "professional competence 专业能力"),
("initiative", "n.", "主动性；倡议", "take the initiative 采取主动"),
("incentive", "n.", "激励，动机", "a financial incentive 经济激励"),
("productivity", "n.", "生产率，生产力", "boost productivity 提高生产率"),
("supervision", "n.", "监督，管理", "under supervision 在监督下"),
("hypothesis", "n.", "假设，假说", "test a hypothesis 检验假设"),
("criterion", "n.", "标准，准则（复 criteria）", "meet the criteria 符合标准"),
("methodology", "n.", "方法论，研究方法", "research methodology 研究方法"),
("implication", "n.", "含义；影响", "far-reaching implications 深远影响"),
("correlation", "n.", "相关性，关联", "a strong correlation 强相关"),
("emission", "n.", "排放（物）", "carbon emissions 碳排放"),
("conservation", "n.", "保护，保存", "wildlife conservation 野生动物保护"),
("ecology", "n.", "生态；生态学", "ecological balance 生态平衡"),
("pollution", "n.", "污染", "air pollution 空气污染"),
("resource", "n.", "资源", "natural resources 自然资源"),
("integrity", "n.", "正直；完整性", "personal integrity 个人操守"),
("dignity", "n.", "尊严，庄重", "human dignity 人的尊严"),
("tolerance", "n.", "宽容，忍耐", "religious tolerance 宗教宽容"),
("sympathy", "n.", "同情，同感", "have sympathy for 同情"),
("gratitude", "n.", "感激，感谢", "express gratitude 表达感激"),
("privilege", "n.", "特权；荣幸", "a privilege to do 做……的荣幸"),
("obstacle", "n.", "障碍，阻碍", "overcome obstacles 克服障碍"),
("burden", "n.", "负担，重担", "a financial burden 经济负担"),
("dimension", "n.", "维度，方面；尺寸", "a new dimension 新层面"),
("boundary", "n.", "边界，界限", "cross the boundary 越过界限"),
("essential", "adj.", "必不可少的；本质的", "essential to health 对健康必需"),
("crucial", "adj.", "至关重要的", "play a crucial role 起关键作用"),
("vital", "adj.", "至关重要的；有活力的", "vital to success 对成功至关重要"),
("fundamental", "adj.", "根本的，基本的", "fundamental principles 基本原则"),
("indispensable", "adj.", "不可或缺的", "an indispensable tool 必不可少的工具"),
("considerable", "adj.", "相当大的，可观的", "a considerable amount 相当大的数量"),
("substantial", "adj.", "大量的；实质的", "substantial progress 实质性进展"),
("excessive", "adj.", "过度的，过分的", "excessive pressure 过大压力"),
("moderate", "adj.", "适度的，温和的", "moderate exercise 适度锻炼"),
("marginal", "adj.", "微小的；边缘的", "a marginal difference 细微差别"),
("reluctant", "adj.", "不情愿的，勉强的", "be reluctant to do 不愿做"),
("skeptical", "adj.", "怀疑的", "be skeptical about 对……怀疑"),
("optimistic", "adj.", "乐观的", "optimistic about the future 对未来乐观"),
("pessimistic", "adj.", "悲观的", "a pessimistic view 悲观看法"),
("indifferent", "adj.", "漠不关心的；平庸的", "indifferent to 对……漠然"),
("ambiguous", "adj.", "模棱两可的，含糊的", "an ambiguous answer 含糊的回答"),
("explicit", "adj.", "明确的，清楚的", "explicit instructions 明确指示"),
("implicit", "adj.", "含蓄的，隐含的", "implicit meaning 隐含意义"),
("inherent", "adj.", "固有的，内在的", "inherent risks 固有风险"),
("inevitable", "adj.", "不可避免的", "an inevitable result 必然结果"),
("capable", "adj.", "有能力的，能干的", "capable of doing 有能力做"),
("competent", "adj.", "胜任的，有能力的", "a competent teacher 称职的老师"),
("eligible", "adj.", "有资格的，合格的", "eligible for the scholarship 有资格申请奖学金"),
("feasible", "adj.", "可行的", "a feasible plan 可行的计划"),
("vulnerable", "adj.", "易受伤害的，脆弱的", "vulnerable to attack 易受攻击"),
("diverse", "adj.", "多样的，不同的", "a diverse culture 多元文化"),
("urban", "adj.", "城市的", "urban development 城市发展"),
("rural", "adj.", "农村的", "rural areas 农村地区"),
("sustainable", "adj.", "可持续的", "sustainable development 可持续发展"),
("contemporary", "adj.", "当代的；同时代的", "contemporary art 当代艺术"),
("nevertheless", "adv.", "然而，不过", "nevertheless, it works 然而它有效"),
("consequently", "adv.", "因此，所以", "consequently, prices rose 因此价格上涨"),
("furthermore", "adv.", "此外，而且", "furthermore, it is cheap 而且它便宜"),
("presumably", "adv.", "大概，想必", "presumably he agreed 想必他同意了"),
("inevitably", "adv.", "不可避免地", "inevitably, costs rise 成本必然上升"),
("emphasize", "v.", "强调，着重", "emphasize the importance 强调重要性"),
("illustrate", "v.", "说明，阐明；加插图", "illustrate the point 说明这一点"),
("demonstrate", "v.", "证明；演示；示威", "demonstrate the ability 展示能力"),
("justify", "v.", "证明……正当", "justify the decision 为决定辩护"),
("verify", "v.", "核实，证实", "verify the data 核实数据"),
("support", "v./n.", "支持；支撑；赡养", "support the proposal 支持提议"),
("oppose", "v.", "反对，抵制", "oppose the plan 反对计划"),
("object", "v.", "反对（object to）", "object to the change 反对改变"),
("approve", "v.", "批准；赞成", "approve of the idea 赞成这个想法"),
("resist", "v.", "抵抗，抵制；忍住", "resist temptation 抵制诱惑"),
("expand", "v.", "扩大，扩展", "expand the business 扩大业务"),
("shrink", "v.", "收缩，缩小；退缩", "shrink in size 尺寸缩小"),
("boost", "v./n.", "促进，提高", "boost confidence 增强信心"),
("decline", "v./n.", "下降；婉拒", "a sharp decline 急剧下降"),
("fluctuate", "v.", "波动，起伏", "prices fluctuate 价格波动"),
("influence", "v./n.", "影响，影响力", "influence the decision 影响决定"),
("affect", "v.", "影响；感动", "affect the result 影响结果"),
("impact", "n./v.", "影响，冲击", "have an impact on 对……有影响"),
("transform", "v.", "转变，改造", "transform society 改变社会"),
("undermine", "v.", "削弱，破坏", "undermine confidence 削弱信心"),
("acquire", "v.", "获得；习得", "acquire knowledge 获取知识"),
("obtain", "v.", "获得，取得", "obtain permission 获得许可"),
("abandon", "v.", "放弃；抛弃", "abandon the plan 放弃计划"),
("sacrifice", "v./n.", "牺牲，献出", "sacrifice time 牺牲时间"),
("deprive", "v.", "剥夺，使丧失", "deprive sb. of sth. 剥夺某人某物"),
("candidate", "n.", "候选人；考生", "a strong candidate 有力人选"),
("colleague", "n.", "同事", "a close colleague 亲密同事"),
("opponent", "n.", "对手，反对者", "a political opponent 政敌"),
("participant", "n.", "参与者", "an active participant 积极参与者"),
("volunteer", "n./v.", "志愿者；自愿做", "volunteer work 志愿工作"),
("survey", "n./v.", "调查；勘测", "conduct a survey 开展调查"),
("questionnaire", "n.", "问卷", "fill in a questionnaire 填写问卷"),
("statistic", "n.", "统计数字（statistics 统计学）", "official statistics 官方统计"),
("sample", "n.", "样本，样品", "a random sample 随机样本"),
("evidence", "n.", "证据，迹象", "scientific evidence 科学证据"),
("crisis", "n.", "危机，紧要关头", "an economic crisis 经济危机"),
("conflict", "n./v.", "冲突，矛盾", "resolve a conflict 解决冲突"),
("dispute", "n./v.", "争端，争论", "settle a dispute 解决争端"),
("threat", "n.", "威胁，恐吓", "pose a threat 构成威胁"),
("risk", "n./v.", "风险；冒……的危险", "take a risk 冒险"),
("innovation", "n.", "创新，革新", "technological innovation 技术创新"),
("strategy", "n.", "策略，战略", "a marketing strategy 营销策略"),
("reform", "n./v.", "改革，改造", "economic reform 经济改革"),
("transition", "n.", "过渡，转变", "a transition period 过渡期"),
("prospect", "n.", "前景，前途", "job prospects 就业前景"),
("alleviate", "v.", "减轻，缓解", "alleviate pain 减轻痛苦"),
("aggravate", "v.", "加重，恶化", "aggravate the problem 使问题恶化"),
("facilitate", "v.", "促进，使便利", "facilitate communication 促进交流"),
("hinder", "v.", "阻碍，妨碍", "hinder progress 阻碍进步"),
("prolong", "v.", "延长，拖延", "prolong life 延长寿命"),
("anticipate", "v.", "预期，预料", "anticipate problems 预见问题"),
("integrate", "v.", "整合，融入", "integrate into society 融入社会"),
("manipulate", "v.", "操纵，控制", "manipulate public opinion 操纵舆论"),
("negotiate", "v.", "谈判，协商", "negotiate a deal 谈成一笔交易"),
("intervene", "v.", "干预，介入", "intervene in the dispute 介入争端"),
("profound", "adj.", "深刻的，深远的", "a profound impact 深远影响"),
("sophisticated", "adj.", "复杂精密的；老练的", "sophisticated technology 精密技术"),
("pragmatic", "adj.", "务实的，实用主义的", "a pragmatic approach 务实的方法"),
("rational", "adj.", "理性的，合理的", "a rational choice 理性选择"),
("arbitrary", "adj.", "任意的，武断的", "an arbitrary decision 武断的决定"),
("controversial", "adj.", "有争议的", "a controversial issue 争议话题"),
("prevalent", "adj.", "普遍的，流行的", "a prevalent view 普遍观点"),
("rigorous", "adj.", "严格的，严谨的", "rigorous training 严格训练"),
("transparent", "adj.", "透明的，易懂的", "a transparent process 透明流程"),
("flexible", "adj.", "灵活的，可弯曲的", "a flexible schedule 灵活的安排"),
("perspective", "n.", "视角，观点；透视", "from a global perspective 从全球视角"),
("insight", "n.", "洞察力，深刻见解", "gain insight into 深入了解"),
("notion", "n.", "概念，观念", "the notion that 认为……的观念"),
("paradox", "n.", "悖论，自相矛盾的事", "a strange paradox 奇怪的悖论"),
("hierarchy", "n.", "等级制度，层级", "a social hierarchy 社会等级"),
("capacity", "n.", "能力；容量", "the capacity to learn 学习能力"),
("potential", "n./adj.", "潜力；潜在的", "great potential 巨大潜力"),
("priority", "n.", "优先事项，优先权", "top priority 首要任务"),
("proportion", "n.", "比例，部分", "a large proportion of 很大一部分"),
("scale", "n.", "规模；刻度；天平", "on a large scale 大规模地"),
("convert", "v.", "转换，改变（信仰）", "convert into 转化为"),
("modify", "v.", "修改，调整", "modify the plan 修改计划"),
("regulate", "v.", "规范，调节，管理", "regulate the market 规范市场"),
("supervise", "v.", "监督，管理", "supervise the work 监督工作"),
("collaborate", "v.", "合作，协作", "collaborate with 与……合作"),
("extensive", "adj.", "广泛的，大量的", "extensive research 广泛研究"),
("intensive", "adj.", "密集的，集中的", "intensive training 强化训练"),
("overwhelming", "adj.", "压倒性的，巨大的", "an overwhelming majority 绝大多数"),
("deliberate", "adj./v.", "故意的；深思熟虑的；商议", "a deliberate attempt 蓄意之举"),
("spontaneous", "adj.", "自发的，自然的", "a spontaneous reaction 自然反应"),
]

from extra_words import EXTRA

_old = {x[0].strip().lower() for x in W}
_dup = [x[0] for x in EXTRA if x[0].strip().lower() in _old]
assert not _dup, _dup
W = W + EXTRA

import sys


def _arg(name, default=None):
    argv = sys.argv[1:]
    if name in argv and argv.index(name) + 1 < len(argv):
        return argv[argv.index(name) + 1]
    return default


DECK_IPA = {}
_deck = _arg('--deck')
if _deck:
    import csv
    rows = list(csv.DictReader(open(_deck, encoding='utf-8-sig')))

    def _pick(row, *names):
        for n in names:
            for k in row:
                if k and k.strip().lower() == n:
                    return (row[k] or '').strip()
        return ''

    W = []
    for row in rows:
        w = _pick(row, 'word', '单词', '英文')
        m = _pick(row, 'meaning', '释义', '中文', '中文意思')
        if not w or not m:
            continue
        W.append((w, _pick(row, 'pos', '词性') or '—', m, _pick(row, 'phrase', '搭配', '例句')))
        ipa = _pick(row, 'ipa', '音标')
        if ipa:
            DECK_IPA[w] = ipa
    if not W:
        raise SystemExit('deck has no usable rows (need word + meaning)')
    TOTAL_WORDS = len(W)
else:
    TOTAL_WORDS = 400

PER_DAY = int(_arg('--per-day', PER_DAY))
_iv = _arg('--intervals')
if _iv:
    INTERVALS = [int(x) for x in _iv.split(',')]
_s = _arg('--start')
if _s:
    _y, _m, _d = [int(x) for x in _s.split('-')]
    START = date(_y, _m, _d)
_OUT = _arg('--out')

_TOTAL_DAYS = (TOTAL_WORDS + PER_DAY - 1) // PER_DAY
EXAM = START + timedelta(days=_TOTAL_DAYS)
_exam = _arg('--exam')
if _exam:
    _ey, _em, _ed = [int(x) for x in _exam.split('-')]
    EXAM = date(_ey, _em, _ed)

assert len(W) == TOTAL_WORDS, len(W)
# 允许最后一天不足 PER_DAY 个（词库总数不一定整除）

try:
    IPA = json.load(open(os.path.join(HERE, 'ipa.json'), encoding='utf-8'))
except Exception:
    IPA = {}

days = []
for idx, start_i in enumerate(range(0, len(W), PER_DAY)):
    chunk = W[start_i:start_i + PER_DAY]
    entry = {
        "day": idx + 1,
        "date": (START + timedelta(days=idx)).isoformat(),
        "words": [{"w": w, "pos": p, "cn": c, "ph": ph, "ipa": DECK_IPA.get(w) or IPA.get(w, "")} for (w, p, c, ph) in chunk],
    }
    days.append(entry)

plan = {"start": START.isoformat(), "exam": EXAM.isoformat(),
        "intervals": INTERVALS, "perDay": PER_DAY, "days": days}

OUTDIR = _OUT or BASE
with open(os.path.join(OUTDIR, "words.json"), "w", encoding="utf-8") as f:
    json.dump(plan, f, ensure_ascii=False, indent=1)

# ---------------- xlsx ----------------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

TOTAL = len(days)
BLUE = "FF1F4E79"
LIGHT = "FFDDEBF7"
GRAY = "FFF2F2F2"
GREEN = "FFC6EFCE"
YELLOW = "FFFFEB9C"
ORANGE = "FFFCE4D6"
thin = Side(style="thin", color="FFBFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEI = "微软雅黑"

wb = Workbook()

# ---------- Sheet 1 每日打卡 ----------
ws = wb.active
ws.title = "每日打卡"
ws.sheet_view.showGridLines = False
REVIEW_HEADS = ["复习%s\n(前%d天)" % ("①②③④⑤⑥"[k], iv) for k, iv in enumerate(INTERVALS)]
heads = ["第几天", "日期", "星期", "今日新词 10 个"] + REVIEW_HEADS + ["完成", "用时\n(分)", "备注"]
widths = [8, 12, 7, 44] + [27] * len(INTERVALS) + [8, 8, 18]
for i, (h, w) in enumerate(zip(heads, widths), start=1):
    c = ws.cell(row=1, column=i, value=h)
    c.font = Font(name=HEI, size=10, bold=True, color="FFFFFFFF")
    c.fill = PatternFill("solid", fgColor=BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 34

WEEK = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
for idx, d in enumerate(days):
    r = idx + 2
    day_no = d["day"]
    cur = date.fromisoformat(d["date"])
    new_txt = " / ".join(f'{x["w"]}（{x["cn"].split("；")[0].split("，")[0]}）' for x in d["words"])
    row_vals = [day_no, cur, WEEK[cur.weekday()], new_txt]
    for k in range(len(INTERVALS)):
        src = day_no - INTERVALS[k]
        if src >= 1:
            row_vals.append(f'第{src}天：' + " ".join(x["w"] for x in days[src - 1]["words"]))
        else:
            row_vals.append("—")
    row_vals += ["☐", None, None]
    for i, v in enumerate(row_vals, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = Font(name=HEI, size=9)
        c.border = BORDER
        c.alignment = Alignment(vertical="center", wrap_text=(i in (4, 5, 6, 7, 8, 9, 10)),
                                horizontal="center" if i in (1, 3, 5 + len(INTERVALS), 6 + len(INTERVALS)) else "left")
    ws.cell(row=r, column=2).number_format = "m月d日"
    if day_no % 5 == 0:  # 每5天一条分隔色
        for i in range(1, 8 + len(INTERVALS)):
            ws.cell(row=r, column=i).fill = PatternFill("solid", fgColor=GRAY)
    ws.row_dimensions[r].height = 64

last = TOTAL + 1
dv = DataValidation(type="list", formula1='"☐,✔,⏭"', allow_blank=True,
                    showDropDown=False, error="请选择 ☐ / ✔ / ⏭")
ws.add_data_validation(dv)
DONE_COL = get_column_letter(5 + len(INTERVALS))
LAST_COL = get_column_letter(7 + len(INTERVALS))
dv.add(f"{DONE_COL}2:{DONE_COL}{last}")
ws.conditional_formatting.add(f"A2:{LAST_COL}{last}",
    FormulaRule(formula=[f'${DONE_COL}2="✔"'], fill=PatternFill(bgColor=GREEN), stopIfTrue=False))
ws.conditional_formatting.add(f"A2:{LAST_COL}{last}",
    FormulaRule(formula=[f'${DONE_COL}2="⏭"'], fill=PatternFill(bgColor=YELLOW), stopIfTrue=False))
ws.freeze_panes = "E2"
ws.auto_filter.ref = f"A1:{LAST_COL}{last}"

# ---------- Sheet 2 单词总表 ----------
ws2 = wb.create_sheet("单词总表")
ws2.sheet_view.showGridLines = False
h2 = ["序号", "第几天", "日期", "单词", "音标", "词性", "释义", "常用搭配 / 短句", "掌握度"]
w2 = [7, 8, 12, 18, 22, 8, 28, 40, 8]
for i, (h, w) in enumerate(zip(h2, w2), start=1):
    c = ws2.cell(row=1, column=i, value=h)
    c.font = Font(name=HEI, size=10, bold=True, color="FFFFFFFF")
    c.fill = PatternFill("solid", fgColor=BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.row_dimensions[1].height = 26

n = 0
for d in days:
    for x in d["words"]:
        n += 1
        r = n + 1
        vals = [n, d["day"], date.fromisoformat(d["date"]), x["w"], x.get("ipa", ""), x["pos"], x["cn"], x["ph"], "○"]
        for i, v in enumerate(vals, start=1):
            c = ws2.cell(row=r, column=i, value=v)
            c.font = Font(name=HEI, size=10, bold=(i == 4))
            c.border = BORDER
            c.alignment = Alignment(vertical="center",
                                    horizontal="center" if i in (1, 2, 5, 6, 9) else "left")
        ws2.cell(row=r, column=3).number_format = "m月d日"
        if d["day"] % 5 == 0:
            for i in range(1, 10):
                ws2.cell(row=r, column=i).fill = PatternFill("solid", fgColor=GRAY)

dv2 = DataValidation(type="list", formula1='"○,△,√"', allow_blank=True, showDropDown=False)
ws2.add_data_validation(dv2)
dv2.add(f"I2:I{n+1}")
ws2.conditional_formatting.add(f"I2:I{n+1}",
    CellIsRule(operator="equal", formula=['"√"'], fill=PatternFill(bgColor=GREEN), stopIfTrue=False))
ws2.conditional_formatting.add(f"I2:I{n+1}",
    CellIsRule(operator="equal", formula=['"△"'], fill=PatternFill(bgColor=YELLOW), stopIfTrue=False))
ws2.freeze_panes = "A2"
ws2.auto_filter.ref = f"A1:I{n+1}"

# ---------- Sheet 3 统计 ----------
ws3 = wb.create_sheet("统计")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions["A"].width = 22
ws3.column_dimensions["B"].width = 18
ws3.column_dimensions["C"].width = 46
t = ws3.cell(row=1, column=1, value="学习进度总览")
t.font = Font(name=HEI, size=14, bold=True, color=BLUE)
ws3.merge_cells("A1:C1")
ws3.row_dimensions[1].height = 30

rows = [
    ("计划开始", START, "m月d日", "Day 1"),
    ("考试日期（估）", EXAM, "m月d日", "以官方通知为准，可自行修改 B3"),
    ("距离考试", "=B3-TODAY()", "0\" 天\"", "每天刷新"),
    ("计划总天数", TOTAL, "0\" 天\"", "每天 10 个新词"),
    ("已打卡天数", f'=COUNTIF(每日打卡!H2:H{last},"✔")', "0\" 天\"", ""),
    ("整体完成率", "=B6/B5", "0.0%", ""),
    ("已学单词数", "=B6*10", "0", "新词累计"),
    ("已掌握（√）", f'=COUNTIF(单词总表!I2:I{n+1},"√")', "0", "能拼写 + 说出释义"),
    ("半熟（△）", f'=COUNTIF(单词总表!I2:I{n+1},"△")', "0", "有印象，需再背"),
    ("今日应学", "=IF(TODAY()<B2,\"未开始\",IF(TODAY()>=B2+B5,\"计划已结束\",\"第 \"&TEXT(TODAY()-B2+1,\"0\")&\" 天\"))", "@", ""),
]
r = 2
for label, val, fmt, note in rows:
    c1 = ws3.cell(row=r, column=1, value=label)
    c1.font = Font(name=HEI, size=10, color="FF404040")
    c1.alignment = Alignment(vertical="center")
    c2 = ws3.cell(row=r, column=2, value=val)
    c2.font = Font(name=HEI, size=11, bold=True, color=BLUE)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.number_format = fmt
    c2.fill = PatternFill("solid", fgColor=LIGHT)
    c2.border = BORDER
    c3 = ws3.cell(row=r, column=3, value=note)
    c3.font = Font(name=HEI, size=9, color="FF808080")
    c3.alignment = Alignment(vertical="center")
    ws3.row_dimensions[r].height = 22
    r += 1

r += 1
tip = ws3.cell(row=r, column=1, value="使用说明")
tip.font = Font(name=HEI, size=12, bold=True, color=BLUE)
r += 1
tips = [
    "1. 每天在「每日打卡」里完成当天一行：新词 10 个 + 3 组复习（前1/4/15天）。",
    "2. 完成后把「完成」列选成 ✔，当天行会自动变绿；整段跳过选 ⏭。",
    "3. 「单词总表」用于查释义、标记掌握度：○ 陌生 / △ 半熟 / √ 已掌握。",
    "4. 复习比新词更重要：哪怕当天没学新词，也要把复习组过一遍。",
    "5. 每天约 25 分钟：新词 10 分钟，复习 15 分钟（每组 5 分钟），总量不超过 40 词。",
]
for line in tips:
    c = ws3.cell(row=r, column=1, value=line)
    c.font = Font(name=HEI, size=9, color="FF404040")
    c.alignment = Alignment(vertical="center")
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws3.row_dimensions[r].height = 20
    r += 1

ws3.sheet_view.zoomScale = 110
wb.save(os.path.join(OUTDIR, "六级单词6周打卡表.xlsx"))
print("ok", len(days), "days,", n, "words")
