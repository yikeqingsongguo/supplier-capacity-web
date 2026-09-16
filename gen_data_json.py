# -*- coding: utf-8 -*-
"""在途数据看板 · 定时数据生成器（供 GitHub Actions 调用）。

读取仓库内三份权威 Excel（**文件名无需固定，直接用 ERP 导出的原始名即可**）：
  data/采购订单*.xlsx  （在途订单，Sheet1；多份时按「最后上传」取最新，见 find_source）
  data/产能*.xlsx      （产能，Sheet1：供应商/供应商简称/采购负责人/月产能）
  data/周计划*.xlsx    （周计划，Sheet1；**可选**，缺失时输出 weekly: [] 不报错）

输出根目录 data.json（结构与在途数据V2.html 内联 DATA 完全一致），
供前端 tryCloudUpdate() 以「静态文件」方式加载——前端 0 次 GitHub API 请求。

本脚本复刻 _inject_data_v3.py + _patch_cap_dual.py + _inject_spu_v4.py 的逻辑，
周计划部分对齐前端 parseWeeklyXLSX（含「文件名日期 + 31 天 = 待交货日期」口径），
路径全部相对仓库根，可在 Linux(Actions)/Windows 本地运行。
"""
import calendar, io, json, os, re, subprocess, sys, openpyxl
from collections import Counter
from datetime import datetime, timedelta, timezone

# 统一使用北京时间(UTC+8)。
# 不能用 datetime.now()：Actions 运行在 UTC，本地 Windows 是 UTC+8，
# 两个产地会产出相差 8 小时的时间戳，导致前端 generatedAt 比较出现「新数据被判为旧」。
BJ = timezone(timedelta(hours=8))

sys.stdout.reconfigure(encoding='utf-8')
ROOT = os.path.dirname(os.path.abspath(__file__))
OUT  = os.path.join(ROOT, 'data.json')
DEFCAP = 15000

ORDER_PAT, ORDER_FIXED, ORDER_LABEL = '采购订单', '采购订单.xlsx', '订单'
CAP_PAT,   CAP_FIXED,   CAP_LABEL   = '产能',     '产能.xlsx',     '产能'
WK_PAT,    WK_FIXED,    WK_LABEL    = '周计划',   '周计划.xlsx',   '周计划'


def find_source(pattern, fixed_name, label):
    """定位源表 —— **无需把文件名改成固定名**，直接用 ERP 导出的原始文件名即可。

    规则：
      1) 先找 data/，data/ 无候选再回退仓库根目录；
      2) 候选 = 该目录下所有 `{pattern}*.xlsx`（排除 Excel 临时文件 ~$* 与 .bak）；
      3) 排序判据（取最大者）：
         ① git 提交时间 —— 真正反映「哪一份最后上传」，适用任意文件名；
         ② 文件名内嵌 14 位时间戳(YYYYMMDDHHMMSS) —— git 不可用时兜底；
         ③ 文件名。无时间戳的固定名优先级最低。
    """
    dirs = [os.path.join(ROOT, 'data'), ROOT]
    cands, from_root = [], False
    for d in dirs:
        if not os.path.isdir(d):
            continue
        found = []
        for fn in sorted(os.listdir(d)):
            if not fn.startswith(pattern) or not fn.lower().endswith('.xlsx'):
                continue
            if fn.startswith('~$') or fn.endswith('.bak'):
                continue
            found.append(os.path.join(d, fn))
        if found:
            cands, from_root = found, (d == ROOT)
            break
    if not cands:
        raise SystemExit('%s源表缺失：请在 data/ 放入 %s，或任意 %s*.xlsx（可保留原始文件名）'
                         % (label, fixed_name, pattern))

    def git_ctime(p):
        """该文件在 git 中最后一次被提交的时间戳；不可用时返回 0。"""
        try:
            out = subprocess.check_output(
                ['git', 'log', '-1', '--format=%ct', '--',
                 os.path.relpath(p, ROOT).replace('\\', '/')],
                cwd=ROOT, stderr=subprocess.DEVNULL).decode().strip()
            return int(out) if out.isdigit() else 0
        except Exception:
            return 0

    def sort_key(p):
        fn = os.path.basename(p)
        m = re.search(r'(\d{14})', fn)
        return (git_ctime(p), 1 if m else 0, m.group(1) if m else '', fn)

    cands.sort(key=sort_key)
    chosen = cands[-1]
    rel = lambda p: os.path.relpath(p, ROOT).replace('\\', '/')
    print('[%s] 源表候选 %d 个: %s' % (label, len(cands), ', '.join(rel(c) for c in cands)))
    print('[%s] 选用: %s' % (label, rel(chosen)))
    if len(cands) > 1:
        print('[%s] 提示: 存在多份源表，已按「最后上传时间」取最新；建议只保留这一份，避免混淆' % label)
    if from_root:
        print('[%s] 提示: 文件位于仓库根目录，建议移动到 data/ 统一管理' % label)
    return chosen

# ---------- 工具 ----------
def num(v):
    if v is None:
        return 0
    s = str(v).strip()
    if s in ('', 'None', 'nan'):
        return 0
    try:
        return int(round(float(s)))
    except Exception:
        return 0

def fmt_date(v):
    if v is None:
        return '—'
    if isinstance(v, datetime):
        return f"{v.year}-{v.month}-{v.day}"
    s = str(v).strip()
    if not s or s == 'None':
        return '—'
    m = re.search(r'(\d{4})\D(\d{1,2})\D(\d{1,2})', s)
    if not m:
        return '—'
    return f"{int(m.group(1))}-{int(m.group(2))}-{int(m.group(3))}"

def tag(v):
    if v is None:
        return ''
    s = str(v).strip()
    return s if s not in ('', 'None') else ''

def spu_of(code):
    if not code:
        return ''
    d = ''.join(ch for ch in str(code) if ch.isdigit())
    return d[:4]

# ---------- 周计划工具（口径对齐前端 parseWeeklyXLSX） ----------
# 前端 four() 取的是「第一个连续 4 位数字」而非「所有数字的前 4 位」，两者在真实数据上
# 结果一致，但为与页面手动导入严格同口径，周计划侧单独实现 four_of()。
_WK_FIELDS = [
    ('materialCode', ('sku', '物料编码', '物料号', '料号')),
    ('materialName', ('物料名称', '品名', '名称', '物料名')),
    ('spec',         ('规格型号', '规格', '型号')),
    ('qty',          ('数量',)),
    ('owner',        ('采购负责人', '负责人')),
    ('supplier',     ('供应商', '供货商', '供应厂商')),
    ('opNote',       ('运营备注', '运营')),
    ('purNote',      ('采购备注', '采购')),
]
_WK_NF = [(k, set(re.sub(r'[　\xa0\s]', '', x).lower() for x in labels))
          for k, labels in _WK_FIELDS]


def _wk_norm(s):
    if s is None:
        return ''
    return re.sub(r'[　\xa0\s]', '', str(s)).lower()


def wk_header_map(rows):
    """在前 40 行内定位表头行：取命中列数最多的一行。返回 (行索引, {字段: 列索引})。"""
    hi, hm, best = -1, {}, 0
    for i in range(min(len(rows), 40)):
        row = rows[i] or []
        m, seen, hit = {}, set(), 0
        for ci in range(len(row)):
            c = _wk_norm(row[ci])
            if not c:
                continue
            for k, labels in _WK_NF:
                if k in seen:
                    continue
                if c in labels:
                    m[k] = ci
                    seen.add(k)
                    hit += 1
                    break
        if hit > best:
            best, hi, hm = hit, i, m
        if hit >= 5:
            break
    return hi, hm


def four_of(v):
    """复刻前端 four()：String(v).match(/\\d{4}/) —— 第一个连续 4 位数字。"""
    if v is None:
        return ''
    m = re.search(r'\d{4}', str(v))
    return m.group(0) if m else ''


def wknum(v):
    """数量：整数原样返回 int（避免 JSON 出现 300.0），非整数保留 4 位。"""
    if v is None:
        return 0
    s = str(v).strip()
    if s in ('', 'None', 'nan'):
        return 0
    try:
        f = float(s)
    except Exception:
        return 0
    return int(round(f)) if abs(f - round(f)) < 1e-9 else round(f, 4)


def month_from_name(n):
    """复刻前端 monthFromName：文件名里第一个 1~12 的数字 → 'N月'，否则 '待排月'。"""
    for x in re.findall(r'\d{1,2}', str(n or '')):
        v = int(x)
        if 1 <= v <= 12:
            return '%d月' % v
    return '待排月'


def plan_date_from_name(name):
    """复刻前端 planDateFromName —— 从文件名解析 ERP 导出日期。

    优先级：① 14 位时间戳(YYYYMMDD) ② 完整日期(YYYY-M-D) ③ 月.日(补当年)。
    解析失败或日期非法返回 ''。返回 'YYYY-MM-DD'。
    """
    s = str(name or '')
    y = m = d = None
    mm = re.search(r'(20\d{2})(\d{2})(\d{2})(?:_\d{4})?', s)
    if mm:
        y, m, d = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))
    else:
        mm = re.search(r'(20\d{2})\s*[-./年]\s*(\d{1,2})\s*[-./月]\s*(\d{1,2})', s)
        if mm:
            y, m, d = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))
        else:
            mm = re.search(r'(\d{1,2})\s*[.月]\s*(\d{1,2})', s)
            if mm:
                y, m, d = datetime.now(BJ).year, int(mm.group(1)), int(mm.group(2))
    if y is None or not (m and 1 <= m <= 12):
        return ''
    if not (d and 1 <= d <= calendar.monthrange(y, m)[1]):
        return ''
    return '%04d-%02d-%02d' % (y, m, d)


def _sniff_weekly(path):
    """表头嗅探：命中列数（必须含 supplier+qty），不合格返回 -1。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sn = 'Sheet1' if 'Sheet1' in wb.sheetnames else wb.sheetnames[0]
        rows = []
        for i, r in enumerate(wb[sn].iter_rows(values_only=True)):
            rows.append(r)
            if i >= 39:
                break
        wb.close()
    except Exception:
        return -1
    hi, hm = wk_header_map(rows)
    return len(hm) if ('supplier' in hm and 'qty' in hm) else -1


def find_weekly_source():
    """周计划源表定位（**可选**，找不到返回 None，不抛错）。

    ① 优先按文件名前缀「周计划」扫 data/ → 仓库根（复用 find_source 的「最后上传」排序）；
    ② 兜底：表头嗅探 —— 扫 data/ 与仓库根所有 xlsx，取表头含「供应商」+「数量」且命中列最多者。
       这一步兼容 ERP 原始文件名（如 9.16tk.xlsx —— 文件名不含「周计划」字样）。
    """
    try:
        return find_source(WK_PAT, WK_FIXED, WK_LABEL)
    except SystemExit:
        pass
    cands = []
    for d in (os.path.join(ROOT, 'data'), ROOT):
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if not fn.lower().endswith('.xlsx') or fn.startswith('~$') or fn.endswith('.bak'):
                continue
            p = os.path.join(d, fn)
            score = _sniff_weekly(p)
            if score > 0:
                cands.append((score, os.path.getmtime(p), p))
    if not cands:
        print('[%s] 未找到源表 —— 本次输出 weekly: []（页面保留内联/手动导入的周计划）' % WK_LABEL)
        return None
    cands.sort(key=lambda x: (x[0], x[1]))
    chosen = cands[-1][2]
    print('[%s] 兜底嗅探选中: %s（表头命中 %d 列）'
          % (WK_LABEL, os.path.relpath(chosen, ROOT).replace('\\', '/'), cands[-1][0]))
    return chosen


def load_weekly(wk_xlsx):
    """解析周计划表 → (weekly 记录列表, 元信息 dict)，结构与前端 DATA.weekly 完全一致。

    口径（已与前端 parseWeeklyXLSX 逐条对齐）：
      · 表头在前 40 行内智能定位，必须命中「供应商」+「数量」
      · 按 (供应商, 物料编码) 分组累加数量；物料名称取最短的那条
      · 供应商名含「国际站」的行视为非供应商，排除（计入 excluded）
      · **文件名日期 + 31 天 = 待交货日期**，据此归属月份；文件名无日期则回退 monthFromName
    """
    fname = os.path.basename(wk_xlsx)
    wb = openpyxl.load_workbook(wk_xlsx, data_only=True, read_only=True)
    sn = 'Sheet1' if 'Sheet1' in wb.sheetnames else wb.sheetnames[0]
    rows = list(wb[sn].iter_rows(values_only=True))
    wb.close()

    hi, hm = wk_header_map(rows)
    if hi < 0 or 'supplier' not in hm or 'qty' not in hm:
        raise SystemExit('周计划表头不合格：需含「供应商」「数量」列'
                         '（可另含 SKU/物料编码/物料名称/规格型号/采购负责人/运营备注/采购备注）')

    def cell(row, k):
        c = hm.get(k)
        return row[c] if (c is not None and c < len(row)) else None

    plan_date = plan_date_from_name(fname)
    if plan_date:
        dt = datetime.strptime(plan_date, '%Y-%m-%d') + timedelta(days=31)
        pre_iso = dt.strftime('%Y-%m-%d')
        pre_date = '%d月%d日' % (dt.month, dt.day)
        pre_month = '%d月' % dt.month
    else:
        pre_iso, pre_date = '', '—'
        pre_month = month_from_name(fname)

    groups, excl = {}, 0
    for i in range(hi + 1, len(rows)):
        row = rows[i] or []
        sup = ('' if cell(row, 'supplier') is None else str(cell(row, 'supplier'))).strip()
        if not sup:
            continue
        if '国际站' in sup:
            excl += 1
            continue
        code = ('' if cell(row, 'materialCode') is None else str(cell(row, 'materialCode'))).strip()
        k = sup + '\x00' + code
        if k not in groups:
            groups[k] = {
                'supplier': sup, 'spu': four_of(code), 'materialCode': code,
                'materialName': ('' if cell(row, 'materialName') is None
                                 else str(cell(row, 'materialName')).strip()),
                'spec': tag(cell(row, 'spec')), 'owner': tag(cell(row, 'owner')),
                'opNote': tag(cell(row, 'opNote')), 'purNote': tag(cell(row, 'purNote')),
                'purchase': 0,
            }
        e = groups[k]
        e['purchase'] += wknum(cell(row, 'qty'))
        mn = ('' if cell(row, 'materialName') is None else str(cell(row, 'materialName')).strip())
        if mn and len(mn) < (len(e['materialName']) or 999):
            e['materialName'] = mn

    out = [{
        'supplier': e['supplier'], 'month': pre_month, 'planDate': plan_date,
        'preDeliveryDate': pre_date, 'preDeliveryIso': pre_iso,
        'received': 0, 'remaining': e['purchase'], 'purchase': e['purchase'],
        'spu': e['spu'], 'materialCode': e['materialCode'], 'materialName': e['materialName'],
        'spec': e['spec'], 'owner': e['owner'], 'opNote': e['opNote'], 'purNote': e['purNote'],
        'deliveryDate': '', 'poNo': '', 'lineTag': '', 'timeTag': '',
        'docType': '周计划', 'creator': '', 'kind': 'weekly', 'status': '待下单', 'bizClose': '',
    } for e in groups.values()]

    info = {
        'file': fname, 'sheet': sn, 'planDate': plan_date, 'preDeliveryIso': pre_iso,
        'preDeliveryDate': pre_date, 'month': pre_month, 'excluded': excl,
        'rows': max(0, len(rows) - hi - 1),
        'suppliers': len(set(e['supplier'] for e in groups.values())),
    }
    return out, info

# ---------- 1) 产能（权威源） ----------
def load_capacity(cap_xlsx):
    """返回 {供应商全名: {'short','owner','effCap'}}，月产能为数字>0 才取值，否则 0。"""
    cap = {}
    wb = openpyxl.load_workbook(cap_xlsx, data_only=True, read_only=True)
    ws = wb.active
    for r in ws.iter_rows(values_only=True):
        sup = (r[0] or '').strip() if r and r[0] is not None else ''
        if not sup or sup == '供应商':
            continue
        raw = r[3] if len(r) > 3 else None
        val = 0
        if isinstance(raw, (int, float)) and raw == raw:
            val = int(raw)
        else:
            m = re.search(r'\d+', str(raw or ''))
            if m:
                val = int(m.group(0))
        cap[sup] = {
            'short': str(r[1]).strip() if len(r) > 1 and r[1] is not None else '',
            'owner': str(r[2]).strip() if len(r) > 2 and r[2] is not None else '',
            'effCap': val if val > 0 else 0,
        }
    return cap

def main():
    # ---------- 0) 源表定位（免改名，自动取最新） ----------
    order_xlsx = find_source(ORDER_PAT, ORDER_FIXED, ORDER_LABEL)
    cap_xlsx   = find_source(CAP_PAT,   CAP_FIXED,   CAP_LABEL)
    wk_xlsx    = find_weekly_source()   # 可选：找不到返回 None，不影响流水线

    CAP = load_capacity(cap_xlsx)

    # ---------- 2) 订单 ----------
    wb = openpyxl.load_workbook(order_xlsx, data_only=True, read_only=True)
    ws = wb['Sheet1']
    rows_iter = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else '' for h in next(rows_iter)]
    idx = {name: i for i, name in enumerate(header)}

    def col(name, *alts):
        if name in idx:
            return idx[name]
        for a in alts:
            if a in idx:
                return idx[a]
        raise SystemExit(f'订单表缺少必需列: {name}（已有表头: {header}）')

    C_PO   = col('单据编号')
    C_BIZ  = col('业务关闭')
    C_SUP  = col('供应商')
    C_CODE = col('物料编码')
    C_DD   = col('交货日期')
    C_REC  = col('累计收料数量')
    C_REM  = col('剩余收料数量')
    C_PUR  = col('采购数量')
    C_NAME = col('物料名称')
    C_LINE = col('产线标签')
    C_TIME = col('时效标签')
    C_CRE  = col('创建人')
    C_DOC  = col('单据类型')

    raw = list(rows_iter)
    po_creator, po_doctype = {}, {}
    for r in raw:
        po_no = (r[C_PO] or '').strip()
        if po_no == '':
            continue
        c = tag(r[C_CRE]); d = tag(r[C_DOC])
        if c:
            po_creator.setdefault(po_no, Counter())[c] += 1
        if d:
            po_doctype.setdefault(po_no, Counter())[d] += 1

    po = []
    months_set, sup_set = set(), set()
    dropped_po = dropped_closed = 0
    for r in raw:
        po_no = (r[C_PO] or '').strip()
        biz = str(r[C_BIZ]).strip() if r[C_BIZ] is not None else ''
        if po_no == '':
            dropped_po += 1
            continue
        if biz == '业务关闭':
            dropped_closed += 1
            continue
        supplier = str(r[C_SUP]).strip()
        material_code = str(r[C_CODE]).strip()
        dd = fmt_date(r[C_DD])
        month = ('—' if dd == '—' else f"{dd.split('-')[1]}月")
        creator = po_creator.get(po_no)
        creator = creator.most_common(1)[0][0] if creator else ''
        doctype = po_doctype.get(po_no)
        doctype = doctype.most_common(1)[0][0] if doctype else ''
        po.append({
            'supplier': supplier,
            'month': month,
            'received': num(r[C_REC]),
            'remaining': num(r[C_REM]),
            'purchase': num(r[C_PUR]),
            'materialCode': material_code,
            'materialName': str(r[C_NAME]).strip(),
            'deliveryDate': dd,
            'fourDigit': material_code[-4:],
            'spu': spu_of(material_code),
            'poNo': po_no,
            'lineTag': tag(r[C_LINE]),
            'timeTag': tag(r[C_TIME]),
            'creator': creator,
            'docType': doctype,
        })
        if month != '—':
            months_set.add(month)
        sup_set.add(supplier)

    # ---------- 2.5) 周计划（可选源表；缺失时输出空数组，不阻断流水线） ----------
    weekly, wk_info = [], None
    if wk_xlsx:
        weekly, wk_info = load_weekly(wk_xlsx)
        for r in weekly:
            if r['month'] and r['month'] != '待排月':
                months_set.add(r['month'])   # 与前端 recompute() 一致：周计划月份并入 months

    months = sorted(months_set, key=lambda m: int(re.sub(r'\D', '', m) or 0))

    # ---------- 3) 产能 / 供应商列表 ----------
    suppliers, capacity = [], []
    for s in sorted(sup_set):
        info = CAP.get(s)
        if info:
            short, owner, eff = info['short'], info['owner'], info['effCap']
            cap_missing = False
            sup_cap = eff
        else:
            short, owner, eff = '', '', 0
            cap_missing = True
            sup_cap = DEFCAP  # suppliers 列表回退占位（与现有管线一致）
        suppliers.append({'name': s, 'short': short, 'owner': owner,
                           'cap': sup_cap, 'capMissing': cap_missing, 'hasPO': True})
        capacity.append({'supplier': s, 'short': short, 'owner': owner,
                          'cap': eff, 'effCap': eff, 'factoryCap': 0})

    real_count = sum(1 for v in CAP.values() if v['effCap'] > 0)
    pending_count = sum(1 for v in CAP.values() if v['effCap'] == 0)

    wk_sup_set = set(r['supplier'] for r in weekly)
    meta = {
        'orderSource': os.path.basename(order_xlsx),
        'orderSheet': 'Sheet1',
        'capSource': '产能表 产能.xlsx (权威源, %d 家真实月产能; 暂定 %d 家按 %d 占位)' % (real_count, pending_count, DEFCAP),
        'rows': len(po), 'poRows': len(po), 'suppliers': len(suppliers),
        'months': months, 'defCap': DEFCAP,
        'generatedAt': datetime.now(BJ).strftime('%Y-%m-%dT%H:%M:%S'),
        'note': '月份由交货日期推导; 数据清洗: 删除空单据编号与业务关闭行; 产线/时效/创建人/单据类型标签来自源表(空则留空); 创建人/单据类型按采购订单编号关联(同PO取最多值); 月产能缺失按默认值补齐',
        'weeklySource':    wk_info['file'] if wk_info else '',
        'weeklySheet':     wk_info['sheet'] if wk_info else '',
        'weeklyRows':      len(weekly),
        'weeklySuppliers': len(wk_sup_set),
        'weeklyPlanDate':  wk_info['planDate'] if wk_info else '',
        'weeklyPreDelivery': wk_info['preDeliveryIso'] if wk_info else '',
        'weeklyMonth':     wk_info['month'] if wk_info else '',
        'weeklyNote': '周计划归属日期 = 源表文件名日期 + 31 天; 口径与前端手动导入完全一致',
    }
    DATA = {'meta': meta, 'suppliers': suppliers, 'months': months,
            'capacity': capacity, 'po': po, 'weekly': weekly}

    with io.open(OUT, 'w', encoding='utf-8') as f:
        json.dump(DATA, f, ensure_ascii=False, separators=(',', ':'))
    print('RAW:', len(raw), '| dropped empty PO:', dropped_po, '| dropped 业务关闭:', dropped_closed)
    print('KEPT:', len(po), '| suppliers:', len(suppliers), '| months:', months)
    print('产能 real/pending:', real_count, pending_count)
    if wk_info:
        print('周计划:', wk_info['file'], '| 源表明细行:', wk_info['rows'], '| 记录:', len(weekly),
              '| 供应商:', wk_info['suppliers'], '| 排除含"国际站":', wk_info['excluded'])
        print('周计划口径: 文件名日期 %s + 31 天 = %s → 归属 %s'
              % (wk_info['planDate'] or '(解析失败)', wk_info['preDeliveryIso'] or '—', wk_info['month']))
    else:
        print('周计划: 未提供源表 → weekly = []（页面保留内联/手动导入的周计划）')
    print('data.json ->', OUT, '(%d bytes)' % os.path.getsize(OUT))

if __name__ == '__main__':
    main()
