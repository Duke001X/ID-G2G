# -*- coding: utf-8 -*-
"""
update_data.py
G2G Indonesia LGM Creative Analytics Dashboard - 数据更新脚本

用法:
    python3 update_data.py LGM数据.xlsx

输入: LGM数据.xlsx，需包含 3 个 sheet:
    - V2L RAW   (短视频引流V2L明细，逐条创意 x 逐天)
    - LGM       (直播广告花费明细，逐条campaign x 逐天)
    - V2L匹配   (创意任务分单表，含 产品/主题/下发账号 等)

输出: data.json (放到仓库根目录，和 index.html 同级)

生成的 data.json 结构:
    lgm          : 按账号+日期聚合的 直播总量 / V2L量 / 话费(hua=直播减V2L)
    creators     : 按达人(下发账号)+日期聚合的 直播数据(从LGM表 Campaign name 解析)
    creator_list : UI筛选chips用的达人白名单(手动维护，见 CREATOR_MERGE_MAP/CREATOR_JUNK)
    records      : 按 Post ID 聚合的创意维度数据(V2L RAW聚合 + V2L匹配表关联prod/theme)
    ana_ct       : records 按 内容类型(ct)+emoji(eg)+账号 二次聚合
    ana_prod     : records 按 产品(prod)+emoji(eg)+账号 二次聚合
    meta         : {updated: 最新日期, min_date: 最早日期}

======================== 核心公式(已用历史数据逐条验证) ========================

--- lgm (按 ACC+DATE) ---
lgm_cost/rev/orders/lv  = LGM表 按账号+日期 sum(Cost/Gross revenue/SKU orders/LIVE views)
lgm_roi                 = lgm_rev / lgm_cost

v2l_cost/rev/orders/lv  = V2L RAW表 按账号+日期 sum(同名列)
v2l_imp/vv/n            = sum(LIVE ads impressions) / sum(Video views) / count(rows)
v2l_ctr                 = sum(LIVE ads clicks) / sum(LIVE ads impressions) * 100
v2l_entry               = sum(LIVE views) / sum(Video views) * 100      # 短视频引流进入直播间的比例
v2l_ret10               = sum(10-second LIVE views) / sum(LIVE views) * 100
v2l_cpm                 = sum(Cost) / sum(LIVE ads impressions) * 1000
v2l_roi                 = v2l_rev / v2l_cost

hua_cost/rev/lv         = lgm_* - v2l_*   (代表除V2L引流之外，直播间本身的花费/收入/观看)
hua_roi                 = hua_rev / hua_cost
v2l_cost_pct            = v2l_cost / lgm_cost * 100
v2l_lv_pct              = v2l_lv / lgm_lv * 100

--- creators (按 达人+DATE，来自 LGM 表) ---
LGM表 Campaign name 格式:
    "Livestream-{达人账号}-{日期/标签}-G"          (常见于4月下旬起，规范格式)
    "Livestream-达播-{达人账号}-{日期}"            (也常见)
只有以 "Livestream-" 开头的 campaign 才算达人直播；
去掉 "Livestream-" 前缀后，如果下一段是 "达播" 则跳过它，取下一段(去掉前导@)作为原始账号名。

原始账号名需要经过 CREATOR_MERGE_MAP / CREATOR_JUNK 归一化:
    - CREATOR_MERGE_MAP: 已知的拼写变体/子号，合并到一个canonical名字下
      (例如 karaskin + karaskinbeauty -> karaskin；pretty.withcaa + pretty withcaa -> pretty.withcaa)
    - CREATOR_JUNK: 3月份的占位/测试代号(IG/K/MIA/id/idd/idn/official/Ashanty等)，归入 "Other"
    - 不在以上两个表里的新名字，默认保留原名(可能是新达人，需要人工看一眼是否要加入白名单)

按 (canonical达人, 日期) 聚合 Cost/Gross revenue/SKU orders/LIVE views，roi=rev/cost。

creator_list: 手动维护的白名单(用于UI筛选按钮)，本脚本默认沿用已有 data.json 里的 creator_list，
不会自动增删。如果有重要新达人(参考脚本输出的 "新增达人" 提示)，需要手动加进去。

--- records (按 Post ID 聚合，来自 V2L RAW + V2L匹配) ---
imp/cost/rev/orders/vv/lv = V2L RAW 按 Post ID sum(对应列)
roi   = rev / cost
ctr   = sum(LIVE ads clicks) / sum(LIVE ads impressions) * 100
cvr   = sum(SKU orders) / sum(LIVE ads clicks) * 100
entry = sum(LIVE views) / sum(Video views) * 100
ret10 = sum(10-second LIVE views) / sum(LIVE views) * 100
ret2s = sum(2-second video views) / sum(Video views) * 100
ret25 = sum(25% video views) / sum(Video views) * 100
cpm   = sum(Cost) / sum(LIVE ads impressions) * 1000
fd/ld/dates = 该 Post ID 出现过的最早/最晚/全部日期

prod/theme: 通过 V2L匹配表的 "Tiktok链接" 提取视频ID，与 Post ID 匹配，
            取该行的 "产品" 作为 prod；"主题" 经 THEME_MAP 归一化后作为 theme。
            匹配不到的 Post ID，prod/theme 为 None。

ct (内容类型) / eg (emoji标签):
    先看 V2L RAW 该 Post ID 的多数 "素材类型":
      - PUGC / Livestream Video / Highlights in live broadcast room /
        Script Library in Live broadcast room / 国内模板 / AI模板库 / 种草主题库 / KOL再剪辑
        -> 直接映射为对应英文 ct (见 CT_MAP)，eg = 标题里检测到的emoji(见下)，没有则 'Other'
      - '#N/A' (即上述几类都不是) 时，进一步判断:
          a) 若该Post ID多数行 "Video source" == "Auto-remix"  -> ct='AI', eg='AI'
          b) 否则，检测标题(Creative列)里是否包含以下emoji之一（按顺序检测第一个命中的）:
                🍀 ❤️ 🔥 💅 💃
             命中 -> ct = 'Untyped ' + 该emoji, eg = 该emoji
          c) 都没命中(既非AI，也没有已知emoji) -> ct='Official', eg='Official'
             (代表官号自然发布的内容，不是走任务分单的素材)

    emoji 分组语义(前端UI用，2026-08 起新增💃):
        直播间(Live Room)  : 🍀 + ❤️
        短视频(Short Video): 🔥 + 💅 + 💃   <- 💃(跳舞)为2026-08新增类别

--- ana_ct / ana_prod ---
在 records 基础上，直接从 V2L RAW 原始数据按 (ct,eg,acc) 或 (prod,eg,acc) 重新聚合(不是从
已四舍五入的records里回算，避免精度损失)，字段:
    cost/rev/imp/vv/lv (sum), n=去重Post ID数量,
    roi=rev/cost, ctr=sum(clicks)/sum(imp)*100, cpm=cost/imp*1000,
    entry=lv/vv*100, ret10=sum(10s live views)/lv*100
==============================================================================
"""

import sys, re, json
from collections import defaultdict, Counter
import openpyxl


# ---------- 达人名归一化表(creators) ----------
CREATOR_MERGE_MAP = {
    'karaskin': 'karaskin', 'karaskinbeauty': 'karaskin',
    'pretty.withcaa': 'pretty.withcaa', 'pretty withcaa': 'pretty.withcaa',
    'glad2glow.official.id': 'glad2glow.official.id', 'glad2glow_official.id': 'glad2glow.official.id',
    'glad2glow_idd': 'glad2glow_idd', 'gladtoglow idd': 'glad2glow_idd', 'official id': 'glad2glow_idd',
    'gladtoglow idn': 'gladtoglow idn',
    'beautecircle': 'beautecircle',
    'syaquee': 'syaquee', 'faeyza_skincare17': 'faeyza_skincare17',
    'silvidnlorenza': 'silvidnlorenza', 'silvidniorenza': 'silvidnlorenza',
    'iqissyy___': 'iqissyy___', 'mommy.csc': 'mommy.csc', 'maulana_840': 'maulana_840',
    'suka.mcflurryoreo': 'suka.mcflurryoreo', 'reyna': 'reyna', 'reayna': 'reyna',
    'glow with me': 'glow with me',
    'dellapermanaa': 'dellapermanaa', 'mia07024': 'mia07024', 'mia': 'mia07024', 'm i a': 'mia07024',
    'novi_novita49': 'novi_novita49',
    'kalisha': 'kalisha', 'inisyraa': 'inisyraa', 'wandaahiliyah': 'wandaahiliyah',
    'hellora_beauty': 'hellora_beauty', 'yuliyana.86': 'yuliyana.86', 'akunkeduakyomi': 'akunkeduakyomi',
    'angelicamanopo': 'angelicamanopo', 'charminenatalie': 'charminenatalie', 'cheryantoinette': 'cheryantoinette',
    'christinathendeano': 'christinathendeano', 'delpaanjurii': 'delpaanjurii', 'hi.preliaa': 'hi.preliaa',
    'itskincara': 'itskincara', 'itsmekarenss': 'itsmekarenss', 'jocelynleora_': 'jocelynleora_',
    'khemiezu': 'khemiezu', 'koko_pino': 'koko_pino', 'lisagultom': 'lisagultom', 'mellyagustiannn': 'mellyagustiannn',
    'nadiatilem': 'nadiatilem', 'ohheyytiara': 'ohheyytiara', 'olviyan': 'olviyan', 'spillspillana': 'spillspillana',
    'teteh.skincaree': 'teteh.skincaree', 'vanessafranzeline': 'vanessafranzeline',
    'wethesibs.official': 'wethesibs.official',
    'farra': 'Farra', 'alfinadamayanti': 'Alfinadamayanti', 'ikoyunichi': 'ikoyunichi',
}
CREATOR_JUNK = {
    'ig', 'id', 'idd', 'idn', 'k', 'ashanty',
    'beauty with ivi', 'kyomi zanira mahreen', 'louis', 'thalia', 'glad to glow id', 'official',
}

# ---------- records: 内容类型 / 主题 归一化表 ----------
CT_MAP = {
    'PUGC': 'PUGC', 'Livestream Video': 'Livestream Video',
    'Highlights in live broadcast room': 'Highlights in live broadcast room',
    'Script Library in Live broadcast room': 'Script Library in Live broadcast room',
    '国内模板': 'Domestic Templates', 'AI模板库': 'AI Template Library',
    '种草主题库': 'Topic Library', 'KOL再剪辑': 'KOL Re-edit',
}
EMOJIS = ['🍀', '❤️', '🔥', '💅', '💃']  # 检测顺序；💃为2026-08新增

THEME_RULES = [
    ('payday sale', 'Promo/Sale'), ('cuci gudang', 'Promo/Sale'), ('4.4 big sale', 'Promo/Sale'),
    ('promo', 'Promo/Sale'), ('discount', 'Promo/Sale'), ('price drop', 'Promo/Sale'),
    ('usp', 'USP'), ('new launch', 'New Launch'),
    ('buy set is cheaper', 'Price/Value'), ('price', 'Price/Value'),
    ('product info', 'Product Info'), ('product knowledge', 'Product Info'),
    ('skin problem', 'Skin Problem'),
    ('bundle', 'Bundle/Set'), ('set', 'Bundle/Set'),
]


def norm_creator(raw):
    key = raw.lower()
    if key in CREATOR_MERGE_MAP:
        return CREATOR_MERGE_MAP[key]
    if key in CREATOR_JUNK:
        return 'Other'
    return raw  # 未知新名字：保留原样，脚本最后会打印提示


def norm_theme(t):
    if not t:
        return None
    t = str(t).strip().lower()
    for k, v in THEME_RULES:
        if k in t:
            return v
    return 'Other'


def extract_campaign_creator(name):
    """从 LGM表 Campaign name 里解析出原始达人账号名(未归一化)"""
    if not name or not name.startswith('Livestream-'):
        return None
    parts = name.split('-')
    rest = parts[1:]
    if not rest:
        return None
    if rest[0].strip() == '达播':
        rest = rest[1:]
    if not rest:
        return None
    return rest[0].strip().lstrip('@')


def detect_emoji(title):
    if not title:
        return None
    for ch in EMOJIS:
        if ch in title:
            return ch
    return None


def build_lgm(lgm_rows, lgm_idx, v2l_rows, v2l_idx):
    lgm_agg = defaultdict(lambda: {'cost': 0, 'rev': 0, 'orders': 0, 'lv': 0})
    for r in lgm_rows:
        acc, date = r[lgm_idx['ACC']], r[lgm_idx['DATE']]
        if not acc or not date:
            continue
        key = (acc, date.strftime('%Y-%m-%d'))
        a = lgm_agg[key]
        a['cost'] += r[lgm_idx['Cost']] or 0
        a['rev'] += r[lgm_idx['Gross revenue']] or 0
        a['orders'] += r[lgm_idx['SKU orders']] or 0
        a['lv'] += r[lgm_idx['LIVE views']] or 0

    v2l_agg = defaultdict(lambda: {'cost': 0, 'rev': 0, 'orders': 0, 'lv': 0, 'imp': 0, 'vv': 0,
                                    'n': 0, 'clicks': 0, 'ret10': 0})
    for r in v2l_rows:
        acc, date = r[v2l_idx['ACC']], r[v2l_idx['DATE']]
        if not acc or not date:
            continue
        key = (acc, date.strftime('%Y-%m-%d'))
        a = v2l_agg[key]
        a['cost'] += r[v2l_idx['Cost']] or 0
        a['rev'] += r[v2l_idx['Gross revenue']] or 0
        a['orders'] += r[v2l_idx['SKU orders']] or 0
        a['lv'] += r[v2l_idx['LIVE views']] or 0
        a['imp'] += r[v2l_idx['LIVE ads impressions']] or 0
        a['vv'] += r[v2l_idx['Video views']] or 0
        a['n'] += 1
        a['clicks'] += r[v2l_idx['LIVE ads clicks']] or 0
        a['ret10'] += r[v2l_idx['10-second LIVE views']] or 0

    out = []
    for key in set(lgm_agg) | set(v2l_agg):
        acc, date = key
        L = lgm_agg.get(key, {'cost': 0, 'rev': 0, 'orders': 0, 'lv': 0})
        V = v2l_agg.get(key, {'cost': 0, 'rev': 0, 'orders': 0, 'lv': 0, 'imp': 0, 'vv': 0,
                               'n': 0, 'clicks': 0, 'ret10': 0})
        lgm_cost, lgm_rev, lgm_orders, lgm_lv = L['cost'], L['rev'], L['orders'], L['lv']
        lgm_roi = lgm_rev / lgm_cost if lgm_cost else 0
        v2l_cost, v2l_rev, v2l_orders, v2l_lv = V['cost'], V['rev'], V['orders'], V['lv']
        v2l_roi = v2l_rev / v2l_cost if v2l_cost else 0
        v2l_imp, v2l_vv, v2l_n = V['imp'], V['vv'], V['n']
        v2l_ctr = V['clicks'] / v2l_imp * 100 if v2l_imp else 0
        v2l_entry = v2l_lv / v2l_vv * 100 if v2l_vv else 0
        v2l_ret10 = V['ret10'] / v2l_lv * 100 if v2l_lv else 0
        v2l_cpm = v2l_cost / v2l_imp * 1000 if v2l_imp else 0
        hua_cost = lgm_cost - v2l_cost
        hua_rev = lgm_rev - v2l_rev
        hua_lv = lgm_lv - v2l_lv
        hua_roi = hua_rev / hua_cost if hua_cost else 0
        v2l_cost_pct = v2l_cost / lgm_cost * 100 if lgm_cost else 0
        v2l_lv_pct = v2l_lv / lgm_lv * 100 if lgm_lv else 0

        out.append({
            'acc': acc, 'date': date, 'month': date[:7],
            'lgm_cost': round(lgm_cost, 3), 'lgm_rev': round(lgm_rev, 3),
            'lgm_orders': lgm_orders, 'lgm_lv': lgm_lv, 'lgm_roi': round(lgm_roi, 3),
            'v2l_cost': round(v2l_cost, 3), 'v2l_rev': round(v2l_rev, 3),
            'v2l_orders': v2l_orders, 'v2l_lv': v2l_lv, 'v2l_roi': round(v2l_roi, 3),
            'v2l_imp': v2l_imp, 'v2l_vv': v2l_vv, 'v2l_n': v2l_n,
            'v2l_ctr': round(v2l_ctr, 3), 'v2l_entry': round(v2l_entry, 3),
            'v2l_ret10': round(v2l_ret10, 3), 'v2l_cpm': round(v2l_cpm, 3),
            'hua_cost': round(hua_cost, 3), 'hua_rev': round(hua_rev, 3),
            'hua_lv': hua_lv, 'hua_roi': round(hua_roi, 3),
            'v2l_cost_pct': round(v2l_cost_pct, 3), 'v2l_lv_pct': round(v2l_lv_pct, 3),
        })
    return out


def build_creators(lgm_rows, lgm_idx):
    agg = defaultdict(lambda: {'cost': 0, 'rev': 0, 'orders': 0, 'lv': 0})
    unknown_names = set()
    for r in lgm_rows:
        date = r[lgm_idx['DATE']]
        if not date:
            continue
        raw = extract_campaign_creator(r[lgm_idx['Campaign name']])
        if not raw:
            continue
        creator = norm_creator(raw)
        if raw.lower() not in CREATOR_MERGE_MAP and raw.lower() not in CREATOR_JUNK:
            unknown_names.add(raw)
        key = (creator, date.strftime('%Y-%m-%d'))
        a = agg[key]
        a['cost'] += r[lgm_idx['Cost']] or 0
        a['rev'] += r[lgm_idx['Gross revenue']] or 0
        a['orders'] += r[lgm_idx['SKU orders']] or 0
        a['lv'] += r[lgm_idx['LIVE views']] or 0

    out = []
    for (creator, date), v in agg.items():
        roi = v['rev'] / v['cost'] if v['cost'] else 0
        out.append({'creator': creator, 'date': date, 'month': date[:7],
                     'cost': round(v['cost'], 2), 'rev': round(v['rev'], 2),
                     'orders': v['orders'], 'lv': v['lv'], 'roi': round(roi, 3)})
    return out, unknown_names


def build_post_info(match_rows, match_idx):
    post_info = {}
    for r in match_rows:
        link = r[match_idx['Tiktok链接']]
        if not link:
            continue
        m = re.search(r'/video/(\d+)', str(link))
        if not m:
            continue
        pid = m.group(1)
        post_info[pid] = {
            'prod': r[match_idx['产品']],
            'theme': norm_theme(r[match_idx['主题']]),
        }
    return post_info


def build_records(v2l_rows, v2l_idx, post_info):
    posts = defaultdict(list)
    for r in v2l_rows:
        pid = r[v2l_idx['Post ID']]
        if not pid:
            continue
        posts[str(pid)].append(r)

    records = []
    for pid, rws in posts.items():
        acc = rws[0][v2l_idx['ACC']]
        title = rws[0][v2l_idx['Creative']]
        mat_types = Counter(r[v2l_idx['素材类型']] for r in rws)
        ct_raw = mat_types.most_common(1)[0][0]
        vsrc = Counter(r[v2l_idx['Video source']] for r in rws)
        is_autoremix = vsrc.most_common(1)[0][0] == 'Auto-remix'
        emoji = detect_emoji(title)

        if ct_raw == '#N/A':
            if is_autoremix:
                ct, eg = 'AI', 'AI'
            elif emoji:
                eg = emoji
                ct = 'Untyped ' + eg
            else:
                ct, eg = 'Official', 'Official'
        else:
            ct = CT_MAP.get(ct_raw, ct_raw)
            eg = emoji or 'Other'

        def s(col):
            return sum(r[v2l_idx[col]] or 0 for r in rws)

        imp, cost, rev, orders = s('LIVE ads impressions'), s('Cost'), s('Gross revenue'), s('SKU orders')
        clicks, vv, lv = s('LIVE ads clicks'), s('Video views'), s('LIVE views')
        v2s, v25, ret10s = s('2-second video views'), s('25% video views'), s('10-second LIVE views')

        roi = rev / cost if cost else 0
        ctr = clicks / imp * 100 if imp else 0
        cvr = orders / clicks * 100 if clicks else 0
        entry = lv / vv * 100 if vv else 0
        ret10 = ret10s / lv * 100 if lv else 0
        ret2s = v2s / vv * 100 if vv else 0
        ret25 = v25 / vv * 100 if vv else 0
        cpm = cost / imp * 1000 if imp else 0
        dates = sorted(set(r[v2l_idx['DATE']].strftime('%Y-%m-%d') for r in rws if r[v2l_idx['DATE']]))
        info = post_info.get(pid, {})

        records.append({
            'id': pid, 'acc': acc, 'eg': eg, 'ct': ct, 'title': title,
            'imp': imp, 'cost': round(cost, 2), 'rev': round(rev, 2), 'orders': orders,
            'roi': round(roi, 3), 'ctr': round(ctr, 3), 'cvr': round(cvr, 3),
            'entry': round(entry, 3), 'ret10': round(ret10, 3), 'ret2s': round(ret2s, 3),
            'ret25': round(ret25, 3), 'vv': vv, 'lv': lv, 'cpm': round(cpm, 3),
            'fd': dates[0] if dates else None, 'ld': dates[-1] if dates else None,
            'dates': ','.join(dates),
            'prod': info.get('prod'), 'theme': info.get('theme'),
        })
    return records


def build_ana(records, v2l_rows, v2l_idx):
    pid_map = {r['id']: (r['ct'], r['eg'], r['prod'], r['acc']) for r in records}
    ct_agg = defaultdict(lambda: {'cost': 0, 'rev': 0, 'imp': 0, 'vv': 0, 'lv': 0, 'clicks': 0, 'ret10s': 0, 'ids': set()})
    prod_agg = defaultdict(lambda: {'cost': 0, 'rev': 0, 'imp': 0, 'vv': 0, 'lv': 0, 'clicks': 0, 'ret10s': 0, 'ids': set()})

    for r in v2l_rows:
        pid = str(r[v2l_idx['Post ID']]) if r[v2l_idx['Post ID']] else None
        if not pid or pid not in pid_map:
            continue
        ct, eg, prod, acc = pid_map[pid]
        cost, rev = r[v2l_idx['Cost']] or 0, r[v2l_idx['Gross revenue']] or 0
        imp, vv, lv = r[v2l_idx['LIVE ads impressions']] or 0, r[v2l_idx['Video views']] or 0, r[v2l_idx['LIVE views']] or 0
        clicks, ret10s = r[v2l_idx['LIVE ads clicks']] or 0, r[v2l_idx['10-second LIVE views']] or 0

        a = ct_agg[(ct, eg, acc)]
        a['cost'] += cost; a['rev'] += rev; a['imp'] += imp; a['vv'] += vv; a['lv'] += lv
        a['clicks'] += clicks; a['ret10s'] += ret10s; a['ids'].add(pid)

        if prod:
            b = prod_agg[(prod, eg, acc)]
            b['cost'] += cost; b['rev'] += rev; b['imp'] += imp; b['vv'] += vv; b['lv'] += lv
            b['clicks'] += clicks; b['ret10s'] += ret10s; b['ids'].add(pid)

    def build(agg, keyname):
        out = []
        for k, v in agg.items():
            roi = v['rev'] / v['cost'] if v['cost'] else 0
            ctr = v['clicks'] / v['imp'] * 100 if v['imp'] else 0
            cpm = v['cost'] / v['imp'] * 1000 if v['imp'] else 0
            entry = v['lv'] / v['vv'] * 100 if v['vv'] else 0
            ret10 = v['ret10s'] / v['lv'] * 100 if v['lv'] else 0
            out.append({keyname: k[0], 'eg': k[1], 'acc': k[2],
                         'cost': round(v['cost'], 2), 'rev': round(v['rev'], 2),
                         'imp': v['imp'], 'vv': v['vv'], 'lv': v['lv'], 'n': len(v['ids']),
                         'roi': round(roi, 3), 'ctr': round(ctr, 3), 'cpm': round(cpm, 3),
                         'entry': round(entry, 3), 'ret10': round(ret10, 3)})
        return out

    return build(ct_agg, 'ct'), build(prod_agg, 'prod')


def main(xlsx_path, old_data_path=None, out_path='data.json'):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    ws_lgm = wb['LGM']
    lgm_headers = [c.value for c in ws_lgm[1]]
    lgm_idx = {h: i for i, h in enumerate(lgm_headers)}
    lgm_rows = list(ws_lgm.iter_rows(min_row=2, values_only=True))

    ws_v2l = wb['V2L RAW']
    v2l_headers = [c.value for c in ws_v2l[1]]
    v2l_idx = {h: i for i, h in enumerate(v2l_headers)}
    v2l_rows = list(ws_v2l.iter_rows(min_row=2, values_only=True))

    ws_match = wb['V2L匹配']
    match_headers = [c.value for c in ws_match[1]]
    match_idx = {h: i for i, h in enumerate(match_headers)}
    match_rows = list(ws_match.iter_rows(min_row=2, values_only=True))

    print('读取完成: LGM', len(lgm_rows), '行, V2L RAW', len(v2l_rows), '行, V2L匹配', len(match_rows), '行')

    lgm = build_lgm(lgm_rows, lgm_idx, v2l_rows, v2l_idx)
    creators, unknown_names = build_creators(lgm_rows, lgm_idx)
    post_info = build_post_info(match_rows, match_idx)
    records = build_records(v2l_rows, v2l_idx, post_info)
    ana_ct, ana_prod = build_ana(records, v2l_rows, v2l_idx)

    # creator_list: 沿用旧 data.json 的白名单(不自动增删)
    creator_list = []
    if old_data_path:
        try:
            old = json.load(open(old_data_path, encoding='utf-8'))
            creator_list = old.get('creator_list', [])
        except Exception as e:
            print('警告: 无法读取旧 data.json 的 creator_list:', e)

    dates = [r['date'] for r in lgm]
    data = {
        'lgm': lgm, 'creators': creators, 'creator_list': creator_list,
        'records': records, 'ana_ct': ana_ct, 'ana_prod': ana_prod,
        'meta': {'updated': max(dates), 'min_date': min(dates)},
    }

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

    print(f'\n生成完成: {out_path}')
    print(f'  日期范围: {min(dates)} ~ {max(dates)}')
    print(f'  lgm: {len(lgm)} 条, creators: {len(creators)} 条, records: {len(records)} 条')
    print(f'  ana_ct: {len(ana_ct)} 条, ana_prod: {len(ana_prod)} 条')
    if unknown_names:
        print(f'\n⚠️  发现 {len(unknown_names)} 个未在 CREATOR_MERGE_MAP/CREATOR_JUNK 中登记的新达人代号(已按原名保留，未归为Other):')
        for n in sorted(unknown_names):
            print('   -', n)
        print('   如果是正式新达人账号，建议加进 CREATOR_MERGE_MAP，并考虑加入 creator_list。')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python3 update_data.py LGM数据.xlsx [旧data.json路径] [输出路径]')
        sys.exit(1)
    xlsx = sys.argv[1]
    old_json = sys.argv[2] if len(sys.argv) > 2 else 'data.json'
    out = sys.argv[3] if len(sys.argv) > 3 else 'data.json'
    main(xlsx, old_json, out)
