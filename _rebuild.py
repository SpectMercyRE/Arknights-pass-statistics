# -*- coding: utf-8 -*-
"""一次性重建：生成 data.js + 复制头像 + 复制势力图标

已固化 v2.22 全部权威修正（博士 2026-10-13 确认）：
1. 显示名映射：统计表错别字 -> 干员表原名；精英干员全大写 -> 首字母大写
2. 头像源映射：精英干员显示名 -> 卫戍协议版头像文件名
3. 系列名：main 加「明日方舟」前缀；超新星删☆后缀、罗小黑战记、6周年礼盒、
   冬隐归路霜星、愚夜密函W、美术设定集VOL.1
4. type：野趣拾光/嘉年华2026 -> carnival；6周年礼盒 -> special（去掉旧「精英干员」误判）
5. 32.0 补维荻（统计表漏录）
6. year 从旧 data.js 按版本号提取
"""
import openpyxl, re, sys, os, shutil
from collections import Counter
from PIL import Image
from pypinyin import lazy_pinyin

sys.stdout.reconfigure(encoding='utf-8')

BASE = r'C:\Users\Administrator\WorkBuddy\2026-10-12-00-40-20'
TONGJI = r'C:\Users\Administrator\WorkBuddy\2026-10-13-18-33-39\明日方舟通行认证统计.xlsx'
GANYUAN = r'C:\Users\Administrator\WorkBuddy\2026-10-13-17-55-01\prts_data\明日方舟干员表.xlsx'
NEW_TABLE = r'C:\Users\Administrator\WorkBuddy\2026-10-13-17-55-01\prts_data\明日方舟与终末地干员总表_按势力.xlsx'
OLD_DATA = BASE + r'\app\www\data.js'  # 旧版：含 year 与旧 series 名（版本号 key）
OUT_DATA = BASE + r'\arknights-pass\data.js'
AV_DIR = r'C:\Users\Administrator\Desktop\干员头像'
AV_DIR_END = r'C:\Users\Administrator\Desktop\终末地干员头像'
FACTION_DIR = r'C:\Users\Administrator\Desktop\势力图标'
DEST_AV = [
    BASE + r'\arknights-pass\images\avatar',
    BASE + r'\app\www\images\avatar',
    BASE + r'\app\android\app\src\main\assets\public\images\avatar',
]
DEST_FACTION = [
    BASE + r'\arknights-pass\images\faction',
    BASE + r'\app\www\images\faction',
    BASE + r'\app\android\app\src\main\assets\public\images\faction',
]

# ===== 显示名映射：统计表名 -> 显示名（改回原名；精英干员全大写 -> 首字母大写）=====
DISPLAY_MAP = {
    '尤里卡': 'U-Official',
    '乌啸': '乌啾',
    '摘英调香师': '撷英调香师',
    '特蕾西娅': '魔王',
    '阿尔图罗': '塑心',
    '安塞尔': '安赛尔',
    '折栀': '折桠',
    '黛菲恩': '戴菲恩',
    '莱因哈特': '莱恩哈特',   # 音律联觉「梦源之地」错别字，= main 6.0 莱恩哈特
    '杰斯顿·威廉姆斯': '杰斯顿',  # 特殊系列「整合运动」简化名（去掉姓氏）
    # 精英干员（白名单凭证/6周年礼盒，全大写 -> 首字母大写）
    'ACE': 'Ace', 'SCOUT': 'Scout', 'OUTCAST': 'Outcast', 'MANTRA': 'Mantra',
    'MISERY': 'Misery', 'MECHANIST': 'Mechanist', 'RAIDIAN': 'Raidian',
    'PITH': 'Pith', 'SHARP': 'Sharp', 'STORMEYE': 'Stormeye', 'TOUCH': 'Touch', 'TULIP': 'Tulip',
}

# ===== 头像源名映射：显示名 -> 头像源文件名（精英干员是卫戍协议版）=====
AVATAR_SRC_MAP = {
    'Misery': 'Misery(卫戍协议)',
    'Mechanist': 'Mechanist(卫戍协议)',
    'Raidian': 'Raidian(卫戍协议)',
    'Pith': 'Pith(卫戍协议)',
    'Sharp': 'Sharp(卫戍协议)',
    'Stormeye': 'Stormeye(卫戍协议)',
    'Touch': 'Touch(卫戍协议)',
    'Tulip': '郁金香(卫戍协议)',
    # 真言 = Mantra 同一人：精一共用 Mantra 图（真言本体无精一立绘）
    '真言': 'Mantra',
}

# 精二头像源覆盖（默认跟随 AVATAR_SRC_MAP）：真言精二用本体立绘（Mantra 无精二）
AVATAR_SRC_MAP_E2 = {
    '真言': '真言',
}

def map_type(cat, series, box):
    b = str(box)
    if cat == '数字系列': return 'main'
    if cat == '联动系列': return 'collab'
    if cat == '明日方舟:终末地': return 'endfield'
    if cat == '嘉年华系列': return 'carnival'
    # 集成映射（含集成映射：野趣拾光）、嘉年华2026 都归嘉年华系列（统计表把它们错放在音律联觉）
    if '集成映射' in b or '嘉年华2026' in b: return 'carnival'
    if cat == '音律联觉系列': return 'sp'
    if cat == '特殊系列':
        # 只有「罗德岛精英干员系列」是 elite；6周年礼盒虽含精英干员但归 special
        if series == '罗德岛精英干员系列': return 'elite'
        return 'special'
    return 'special'

def clean_name(n):
    return re.sub(r'（仅精一）|（特殊款）', '', n).strip()

# ===== 终末地干员头像：源目录用中点（管理员·男/女），data.js 用全角括号（管理员（男/女））=====
END_AVATAR_MAP = {
    '管理员（男）': '管理员·男',
    '管理员（女）': '管理员·女',
}

# ===== 终末地干员星级（博士 2026-10-14 提供《终末地干员一览.md》，35 名；职业/团队暂无数据）=====
END_OPERATORS = {
    # 六星（20 个形态：管理员男女各算一个）
    '伊冯': 6, '余烬': 6, '别礼': 6, '卡缪': 6, '庄方宜': 6, '弭弗': 6,
    '提弗洛斯': 6, '明河': 6, '梨诺': 6, '汤汤': 6, '洁尔佩塔': 6, '洛茜': 6,
    '祀': 6, '管理员（男）': 6, '管理员（女）': 6, '艾尔黛拉': 6, '莱万汀': 6,
    '诀': 6, '骏卫': 6, '黎风': 6,
    # 五星
    '佩丽卡': 5, '噗切娜': 5, '大潘': 5, '弧光': 5, '昼雪': 5, '狼卫': 5,
    '艾维文娜': 5, '赛希': 5, '阿列什': 5, '陈千语': 5,
    # 四星
    '卡契尔': 4, '埃特拉': 4, '安塔尔': 4, '秋栗': 4, '萤石': 4,
}

# ===== 终末地干员英文名（博士 2026-10-14 提供《终末地干员一览.md》）=====
# 注：管理员·男/女 在 data.js 显示名为「管理员（男）/（女）」（全角括号），共用代号 Endministrator
END_EN_NAMES = {
    '伊冯': 'Yvonne', '余烬': 'Ember', '别礼': 'Last Rite', '卡缪': 'Camille',
    '庄方宜': 'Zhuang Fangyi', '弭弗': 'Mifu', '提弗洛斯': 'Typhoeus', '明河': 'Minghe',
    '梨诺': 'Liino', '汤汤': 'Tangtang', '洁尔佩塔': 'Gilberta', '洛茜': 'Rossi',
    '祀': 'Si', '管理员（男）': 'Endministrator', '管理员（女）': 'Endministrator',
    '艾尔黛拉': 'Ardelia', '莱万汀': 'Laevatain', '诀': 'Arcane', '骏卫': 'Pogranichnik',
    '黎风': 'Lifeng',
    '佩丽卡': 'Perlica', '噗切娜': 'Purrchena', '大潘': 'Da Pan', '弧光': 'Arclight',
    '昼雪': 'Snowshine', '狼卫': 'Wulfgard', '艾维文娜': 'Avywenna', '赛希': 'Xaihi',
    '阿列什': 'Alesh', '陈千语': 'Chen Qianyu',
    '卡契尔': 'Catcher', '埃特拉': 'Estella', '安塔尔': 'Antal', '秋栗': 'Akekuri', '萤石': 'Fluorite',
}

def pinyin_fields(name):
    """干员名 -> (全拼 py, 首字母缩写 pyi)，小写且仅保留字母"""
    seg = lazy_pinyin(name)
    py = re.sub(r'[^a-z]', '', ''.join(seg).lower())
    pyi = re.sub(r'[^a-z]', '', ''.join((s[0] if s else '') for s in seg).lower())
    return py, pyi

def crop_square(img):
    """非正方形图裁成中心正方形"""
    w, h = img.size
    s = min(w, h)
    left = (w - s) // 2
    top = (h - s) // 2
    return img.crop((left, top, left + s, top + s))

def final_series(box, t):
    b = str(box)
    if '超新星' in b:
        return '2024不觅浪尘SP-FIVE超新星'          # 删掉☆后面的字
    if '罗小黑' in b:
        return '通行认证SP-罗小黑战记'
    if '6周年礼盒' in b:
        return '鹰角网络6周年礼盒'
    if '霜星' in b and '礼盒' in b:
        return '冬隐归路 霜星角色礼盒'
    if 'W怪盗' in b:
        return '愚夜密函礼盒-W角色套组'
    if '博士通行证' in b:
        return '明日方舟官方美术设定集VOL.1'
    if t == 'main':
        return '明日方舟' + b                          # "通行认证 X.0" -> "明日方舟通行认证 X.0"
    return b

# ===== 1. 提取 year_map（版本号 -> year）=====
year_map = {}
old_text = open(OLD_DATA, encoding='utf-8').read()
for m in re.finditer(r'\{ series: "([^"]+)", type: "([^"]+)", year: "([^"]*)", characters:', old_text):
    sn, yv = m.group(1), m.group(3)
    if not yv:
        continue
    ver = re.search(r'(\d+)\.0', sn)
    key = ver.group(0) if ver else sn
    year_map[key] = yv
print('year_map 条目:', len(year_map))

# ===== 2. 读干员表 -> OPERATOR_INFO（key=干员表原名）=====
wb2 = openpyxl.load_workbook(GANYUAN, read_only=True, data_only=True)
ws2 = wb2['明日方舟干员表']
op_raw = {}
for r in ws2.iter_rows(values_only=True):
    if r[1] is None:
        continue
    name = str(r[1]).strip()
    if name in ('干员名称',):
        continue
    op_raw[name] = {
        'class': str(r[2]).strip() if r[2] else '',
        'rarity': r[3] if r[3] is not None else '',
        'faction': str(r[4]).strip() if r[4] else '',
        'nation': str(r[5]).strip() if r[5] else '',
    }
wb2.close()

# 终末地干员：星级来自《终末地干员一览.md》（职业暂无，稍后由新表补英文名/势力）
for en, rar in END_OPERATORS.items():
    op_raw[en] = {'class': '', 'rarity': rar, 'faction': '', 'nation': ''}

# ===== 2b. 读新表（按势力）：全干员英文名 + 势力 + 职业 =====
# 列：0序号 1游戏 2干员名称 3职业 4星级 5英文名 6所属势力
new_info = {}
wb3 = openpyxl.load_workbook(NEW_TABLE, read_only=True, data_only=True)
ws3 = wb3['明日方舟干员表']
for r in ws3.iter_rows(values_only=True):
    if r[2] is None:
        continue
    nm = str(r[2]).strip()
    if nm == '干员名称':
        continue
    new_info[nm] = {
        'class': str(r[3]).strip() if r[3] else '',
        'rarity': r[4] if r[4] is not None else '',
        'en': str(r[5]).strip() if r[5] else '',
        'faction': str(r[6]).strip() if r[6] else '',
    }
wb3.close()

# 终末地管理员：新表用中点（管理员·男），data.js 显示名用全角括号（管理员（男））
END_NAME_REV = {'管理员·男': '管理员（男）', '管理员·女': '管理员（女）'}

# 用新表补齐英文名 + 势力 + 职业（方舟本体英文名此前为空；终末地无职业 -> class 空）
for name, info in op_raw.items():
    src = END_AVATAR_MAP.get(name, name)  # 管理员（男）-> 管理员·男
    nw = new_info.get(src)
    if nw:
        info['en'] = nw['en']
        info['faction'] = nw['faction']
        info['class'] = nw['class']

# 新表多出的干员（如大帝）：补入 op_raw
for nm, nw in new_info.items():
    key = END_NAME_REV.get(nm, nm)
    if key not in op_raw:
        op_raw[key] = {
            'class': nw['class'],
            'rarity': nw['rarity'],
            'faction': nw['faction'],
            'nation': '',
            'en': nw['en'],
        }

# ===== 2c. 剧情/NPC 干员补录（不在干员表中，但 PASS_DATA 里有通行证）=====
# 博士指定：霜星/博士/Ace/Scout/Outcast/Mantra 归属罗德岛阵营
NPC_EXTRA = {
    'Ace':     {'en': 'Ace',       'faction': '罗德岛'},
    'Scout':   {'en': 'Scout',     'faction': '罗德岛'},
    'Outcast': {'en': 'Outcast',   'faction': '罗德岛'},
    'Mantra':  {'en': 'Mantra',    'faction': '罗德岛'},
    '霜星':    {'en': 'FrostNova', 'faction': '罗德岛'},
    '博士':    {'en': 'Doctor',    'faction': '罗德岛'},
}
for nm, ex in NPC_EXTRA.items():
    if nm not in op_raw:
        op_raw[nm] = {'class': '', 'rarity': '', 'nation': '', 'py': '', 'pyi': '',
                      'en': ex['en'], 'faction': ex['faction']}
    else:
        op_raw[nm]['faction'] = ex['faction']
        op_raw[nm].setdefault('en', ex['en'])

# 给所有干员补拼音 py / 首字母 pyi（供搜索）
for name, info in op_raw.items():
    py, pyi = pinyin_fields(name)
    info['py'] = py
    info['pyi'] = pyi
    info['en'] = info.get('en', '')

# ===== 3. 读统计表 =====
wb = openpyxl.load_workbook(TONGJI, read_only=True, data_only=True)
ws = wb['通行认证清单']
rows = list(ws.iter_rows(values_only=True))[1:]

pass_data = []
only_e1 = set()
all_char_names = set()
for r in rows:
    if r[0] is None or r[3] is None:
        continue
    cat, series, box = r[1], r[2], r[3]
    t = map_type(cat, series, box)
    raw_chars = [c.strip() for c in str(r[4]).split('、')] if r[4] else []
    chars = []
    for c in raw_chars:
        if not c:
            continue
        if '（仅精一）' in c:
            cn = clean_name(c)
            only_e1.add(DISPLAY_MAP.get(cn, cn))
        cn = clean_name(c)
        dn = DISPLAY_MAP.get(cn, cn)   # 显示名
        chars.append(dn)
        all_char_names.add(dn)
    # 统计表漏录的干员：32.0 补维荻（4星特种）
    if t == 'main' and '32.0' in str(box):
        chars.append('维荻')
        all_char_names.add('维荻')
    year = ''
    ym = re.search(r'(19|20)\d{2}', str(box))
    if ym:
        year = ym.group(0)
    if t == 'main':
        ver = re.search(r'(\d+)\.0', str(box))
        if ver and ver.group(0) in year_map:
            year = year_map[ver.group(0)]
    pass_data.append({'series': final_series(box, t), 'type': t, 'year': year, 'characters': chars})
wb.close()

# ===== 统计表漏录：终末地新增 2 盒（管理员凭证Beta、管理员凭证4期，博士 2026-10-13 确认）=====
pass_data.append({'series': '管理员凭证Beta', 'type': 'endfield', 'year': '', 'characters': ['管理员（女）', '管理员（男）']})
pass_data.append({'series': '管理员凭证4期', 'type': 'endfield', 'year': '', 'characters': ['管理员（女）', '管理员（男）']})

# ===== 排序：按类别分组（数字→音律联觉→嘉年华→联动→精英→特殊→终末地），类别内部按规则排 =====
def sort_key(d):
    t = d['type']
    s = d['series']
    if t == 'main':
        m = re.search(r'(\d+)\.0', s)
        return (0, int(m.group(1)) if m else 0)        # 数字系列从小到大
    if t == 'sp':
        m = re.search(r'(19|20)\d{2}', s)
        return (1, int(m.group(0)) if m else 0)        # 音律联觉按年份
    if t == 'carnival':
        # 嘉年华2022 → 2024 → 2026 → 集成映射 → 集成映射：野趣拾光
        if '嘉年华2022' in s: return (2, 0)
        if '嘉年华2024' in s: return (2, 1)
        if '嘉年华2026' in s: return (2, 2)
        if s == '通行认证SP-集成映射': return (2, 3)
        if '野趣拾光' in s: return (2, 4)
        return (2, 5)
    if t == 'collab': return (3, 0)
    if t == 'elite': return (4, 0)
    if t == 'special': return (5, 0)
    if t == 'endfield':
        # 白名单1.0→2.0→3.0→SP嘉年华→管理员Beta→1期→2期→3期→4期
        if '管理员凭证' in s:
            if 'Beta' in s or 'beta' in s:
                return (6, 5)
            m = re.search(r'(\d+)期', s)
            return (6, 6 + (int(m.group(1)) if m else 0))
        if '白名单凭证' in s:
            if '嘉年华' in s:
                return (6, 4)
            m = re.search(r'(\d+)\.0', s)
            return (6, int(m.group(1)) if m else 0)
        return (6, 99)
    return (99, 0)

pass_data.sort(key=sort_key)

# ===== LOW_STAR：星级<=3（干员表名=显示名，直接取）+ 仅精一 =====
low_star = set()
for raw_name, info in op_raw.items():
    rar = info['rarity']
    if isinstance(rar, (int, float)) and rar <= 3:
        low_star.add(raw_name)
low_star |= only_e1

# ===== 势力 logo 处理：生成黑白双版本（浅色模式用黑、深色模式用白）=====
FACTION_LOGO_SRC = {
    '罗德岛':   '罗德岛.png',
    'Ave Mujica': 'Ave_Mujica.png',
    'S.E.E.S.': 'S.E.E.S.png',
    '炎':       'Logo_炎.png',
    '维多利亚': 'Logo_维多利亚.png',
    '哥伦比亚': 'Logo_哥伦比亚.png',
    '终末地':   'Logo_终末地.png',
    '萨尔贡':   'Logo_萨尔贡.png',
    '叙拉古':   'Logo_叙拉古.png',
    '乌萨斯':   '乌萨斯.png',
    '拉特兰':   'Logo_拉特兰.png',
    '雷姆必拓': 'Logo_雷姆必拓.png',
    '卡西米尔': 'Logo_卡西米尔.png',
    '东':       'Logo_东.png',
    '伊比利亚': 'Logo_伊比利亚.png',
    '莱塔尼亚': 'Logo_莱塔尼亚.png',
    '谢拉格':   'Logo_谢拉格.png',
    '彩虹小队': 'Logo_彩虹小队.png',
    '玻利瓦尔': 'Logo_玻利瓦尔.png',
    '莱茵生命': '莱茵生命.png',
    '企鹅物流': '企鹅物流.png',
    '米诺斯':   'Logo_米诺斯.png',
    '萨米':     'Logo_萨米.png',
    '阿戈尔':   'Logo_阿戈尔.png',
    '深海猎人': '深海猎人.png',
    '黑钢国际': '黑钢.png',
    '巴别塔':   'Logo_巴别塔.png',
    '使徒':     '使徒.png',
    '汐斯塔':   'Logo_汐斯塔.png',
}
# 处理方式：single=单版本保持原色（彩色/自带颜色，深浅共用）；dual=纯色黑白双版本；sees=反色双版本（黑↔白、蓝色保留）
# rainbow=浅色用 mono 黑版（原图黑部分在浅色卡上才看得清）、深色用原版；scale=显示缩放系数（Ave Mujica 偏小调大）
FACTION_LOGO_MODE = {
    '罗德岛': 'single',
    'Ave Mujica': 'single',
    'S.E.E.S.': 'sees',
    '彩虹小队': 'rainbow',
}
FACTION_LOGO_SCALE = {
    'Ave Mujica': 1.25,
}

def logo_key(faction):
    return faction.rstrip('.')

def make_mono(img, color):
    """非透明像素统一成纯色（黑/白），保留 alpha 抗锯齿"""
    a = img.split()[-1]
    solid = Image.new('L', img.size, color)
    return Image.merge('RGBA', (solid, solid, solid, a))

def invert_grayscale(img):
    """灰度像素反色（黑↔白），彩色像素（SEES 蓝）保留"""
    src = img.load()
    out = img.copy()
    dst = out.load()
    w, h = img.size
    for y in range(h):
        for x in range(w):
            r, g, b, a = src[x, y]
            if a == 0:
                continue
            if abs(r - g) < 20 and abs(g - b) < 20 and abs(r - b) < 20:
                dst[x, y] = (255 - r, 255 - g, 255 - b, a)
    return out

def invert_full(img):
    """整图反色（黑↔白），保留 alpha——用于黑白图案类 logo（彩虹小队）：
    mono 单色会抹掉图案内部的黑白对比（反色后白色部分变黑才在浅色卡上可见）"""
    src = img.split()
    r = src[0].point(lambda v: 255 - v)
    g = src[1].point(lambda v: 255 - v)
    b = src[2].point(lambda v: 255 - v)
    return Image.merge('RGBA', (r, g, b, src[3]))

def normalize_logo(img, canvas=96):
    """统一尺寸：裁掉透明边 → 等比缩放到 canvas 内 → 居中放到正方形透明画布上。
    源图像素尺寸五花八门（有的带大片留白、有的顶满），不归一化会导致卡片水印忽大忽小。"""
    bbox = img.getchannel('A').getbbox()
    if bbox:
        img = img.crop(bbox)
    img.thumbnail((canvas, canvas), Image.LANCZOS)
    out = Image.new('RGBA', (canvas, canvas), (0, 0, 0, 0))
    out.paste(img, ((canvas - img.width) // 2, (canvas - img.height) // 2), img)
    return out

# 清空并重建势力图标目录，生成 b（浅色黑版）/ w（深色白版）
for d in DEST_FACTION:
    if os.path.exists(d):
        shutil.rmtree(d)
    os.makedirs(d, exist_ok=True)

faction_logo_map = {}
fcnt = 0
for faction, src_file in FACTION_LOGO_SRC.items():
    mode = FACTION_LOGO_MODE.get(faction, 'dual')
    key = logo_key(faction)
    src_path = os.path.join(FACTION_DIR, src_file)
    if not os.path.exists(src_path):
        print('  [跳过] 缺失源图标:', src_file)
        continue
    img = Image.open(src_path).convert('RGBA')
    img = normalize_logo(img)  # 全部归一化到 96×96 画布，卡片水印大小一致
    b_file = key + '.png'
    if mode == 'single':
        w_file = b_file
        for d in DEST_FACTION:
            img.save(os.path.join(d, b_file), 'PNG')
    elif mode == 'dual':
        w_file = key + '_w.png'
        b_img = make_mono(img, 0)
        w_img = make_mono(img, 255)
        for d in DEST_FACTION:
            b_img.save(os.path.join(d, b_file), 'PNG')
            w_img.save(os.path.join(d, w_file), 'PNG')
    elif mode == 'rainbow':
        # 彩虹小队：黑白图案 logo，浅色模式用整图反色版（白↔黑，图案内部对比保留），
        # 深色模式用原版（白色部分可见）——mono 单色会把图案内部细节糊成一团，弃用
        w_file = key + '_w.png'
        b_img = invert_full(img)
        for d in DEST_FACTION:
            b_img.save(os.path.join(d, b_file), 'PNG')
            img.save(os.path.join(d, w_file), 'PNG')
    else:  # sees：原版（黑+蓝）为浅色，反色（白+蓝）为深色
        w_file = key + '_w.png'
        w_img = invert_grayscale(img)
        for d in DEST_FACTION:
            img.save(os.path.join(d, b_file), 'PNG')
            w_img.save(os.path.join(d, w_file), 'PNG')
    entry = {'b': b_file, 'w': w_file}
    if faction in FACTION_LOGO_SCALE:
        entry['s'] = FACTION_LOGO_SCALE[faction]
    faction_logo_map[faction] = entry
    fcnt += 1

# ===== 写 data.js =====
def js_str(s):
    return '"' + s.replace('\\', '\\\\').replace('"', '\\"') + '"'

lines = []
lines.append('// 明日方舟通行证收集数据（由通行认证统计表自动生成）')
lines.append('// 系列名 = 统计表盒名（main 加「明日方舟」前缀）；type: main数字/sp音律/carnival嘉年华/collab联动/elite精英/special特殊/endfield终末地')
lines.append('// OPERATOR_INFO：干员职业/星级/团队/国家（团队未校对，暂不用于卡片）')
lines.append('// OPERATOR_ALIAS：异名映射（本轮已全部改为原名，留空备用）')
lines.append('')
lines.append('// 三星及以下 + 仅精一干员（只有精一，无精二立绘，通行证只有 1 张）')
lines.append('window.LOW_STAR = new Set([')
low_list = sorted(low_star)
for i in range(0, len(low_list), 6):
    chunk = low_list[i:i+6]
    lines.append('  ' + ', '.join(js_str(x) for x in chunk) + ',')
lines.append(']);')
lines.append('')
lines.append('window.PASS_DATA = [')
type_comment = {
    'main': '========== 数字系列（主线）==========',
    'sp': '========== 音律联觉系列 ==========',
    'carnival': '========== 嘉年华系列 ==========',
    'collab': '========== 联动系列 ==========',
    'elite': '========== 精英干员系列（白名单凭证）==========',
    'special': '========== 特殊系列 ==========',
    'endfield': '========== 明日方舟：终末地 ==========',
}
last_type = None
for d in pass_data:
    if d['type'] != last_type:
        lines.append('')
        lines.append('  // ' + type_comment[d['type']])
        last_type = d['type']
    chars = ', '.join(js_str(c) for c in d['characters'])
    lines.append('  { series: %s, type: %s, year: %s, characters: [%s] },' % (js_str(d['series']), js_str(d['type']), js_str(d['year']), chars))
lines.append('];')
lines.append('')
lines.append('// 干员信息：职业/星级/势力/国家（势力取自按势力总表，供卡片 logo 与搜索）')
lines.append('window.OPERATOR_INFO = {')
for name in sorted(op_raw.keys()):
    info = op_raw[name]
    rar = info['rarity']
    if isinstance(rar, float) and rar == int(rar):
        rar_s = str(int(rar))
    elif rar == '':
        rar_s = '""'
    else:
        rar_s = str(rar)
    lines.append('  %s: { class: %s, rarity: %s, faction: %s, nation: %s, en: %s, py: %s, pyi: %s },' % (
        js_str(name), js_str(info['class']), rar_s, js_str(info['faction']), js_str(info['nation']),
        js_str(info.get('en', '')), js_str(info.get('py', '')), js_str(info.get('pyi', ''))))
lines.append('};')
lines.append('')
lines.append('// 势力 logo 文件映射：{ b: 浅色模式黑版, w: 深色模式白版, s: 显示缩放系数(可选) }（单版本势力 b==w）')
lines.append('window.FACTION_LOGO = {')
for faction in sorted(faction_logo_map.keys()):
    m = faction_logo_map[faction]
    s_part = (', s: %s' % m['s']) if 's' in m else ''
    lines.append('  %s: { b: %s, w: %s%s },' % (js_str(faction), js_str(m['b']), js_str(m['w']), s_part))
lines.append('};')
lines.append('')
lines.append('// 异名映射：统计表干员名 -> 干员表名（本轮已全部改为原名，留空备用）')
lines.append('window.OPERATOR_ALIAS = {')
lines.append('};')
lines.append('')
lines.append('// 需要按「精一/精二」拆分的系列（数字系列 + 联动系列）')
lines.append('window.SPLIT_ELITE_SERIES = new Set([')
lines.append('  ...window.PASS_DATA.filter(s => s.type === \'main\' || s.type === \'collab\').map(s => s.series)')
lines.append(']);')
lines.append('')

with open(OUT_DATA, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

# ===== 复制头像（显示名 -> 头像源文件名）=====
for d in DEST_AV:
    if os.path.exists(d):
        shutil.rmtree(d)
    os.makedirs(d, exist_ok=True)

copied = 0
missing_e1 = []   # 精一缺失（该干员完全无头像）
for name in sorted(all_char_names):
    # 终末地干员：从终末地源目录取，裁成正方形（管理员·男/女 -> 管理员（男/女））
    end_fn = END_AVATAR_MAP.get(name, name)
    end_src = os.path.join(AV_DIR_END, f'{end_fn}.png')
    if os.path.exists(end_src):
        img = Image.open(end_src)
        if img.mode in ('RGBA', 'LA', 'P'):
            img = img.convert('RGBA')
            bg = Image.new('RGB', img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        else:
            img = img.convert('RGB')
        img = crop_square(img)
        dst_name = f'{name}-精一.png'
        for d in DEST_AV:
            img.save(os.path.join(d, dst_name), 'PNG')
        copied += 1
        continue
    # 方舟干员：精一 + 精二
    src_name = AVATAR_SRC_MAP.get(name, name)
    for elite in ('精一', '精二'):
        src_name_e = src_name if elite == '精一' else AVATAR_SRC_MAP_E2.get(name, src_name)
        src = os.path.join(AV_DIR, f'{src_name_e}-{elite}.png')
        if not os.path.exists(src):
            if elite == '精一':
                missing_e1.append(name)
            continue
        dst_name = f'{name}-{elite}.png'
        for d in DEST_AV:
            shutil.copy2(src, os.path.join(d, dst_name))
        copied += 1

# ===== 汇总 =====
print()
print('=== 重建汇总 ===')
print('系列总数:', len(pass_data))
print('分类分布:', dict(Counter(d['type'] for d in pass_data)))
print('干员信息总数:', len(op_raw))
print('低星/仅精一:', len(low_star))
print('复制头像文件:', copied, '个（三处各', copied, '）')
print('生成势力图标:', fcnt, '个势力（各含 b/w 两版）')
print('精一缺失(留空) 干员:', len(missing_e1))
for m in missing_e1:
    print('  -', m)
print()
print('已写出 data.js:', OUT_DATA)
