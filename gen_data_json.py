# -*- coding: utf-8 -*-
"""在途数据看板 · 定时数据生成器（供 GitHub Actions 调用）。

读取仓库内两份权威 Excel（**文件名无需固定，直接用 ERP 导出的原始名即可**）：
  data/采购订单*.xlsx  （在途订单，Sheet1；多份时按文件名时间戳取最新）
  data/产能*.xlsx      （产能，Sheet1：供应商/供应商简称/采购负责人/月产能）

输出根目录 data.json（结构与在途数据V2.html 内联 DATA 完全一致），
供前端 tryCloudUpdate() 以「静态文件」方式加载——前端 0 次 GitHub API 请求。

本脚本复刻 _inject_data_v3.py + _patch_cap_dual.py + _inject_spu_v4.py 的逻辑，
路径全部相对仓库根，可在 Linux(Actions)/Windows 本地运行。
"""
import io, json, os, re, subprocess, sys, openpyxl
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

    meta = {
        'orderSource': os.path.basename(order_xlsx),
        'orderSheet': 'Sheet1',
        'capSource': '产能表 产能.xlsx (权威源, %d 家真实月产能; 暂定 %d 家按 %d 占位)' % (real_count, pending_count, DEFCAP),
        'rows': len(po), 'poRows': len(po), 'suppliers': len(suppliers),
        'months': months, 'defCap': DEFCAP,
        'generatedAt': datetime.now(BJ).strftime('%Y-%m-%dT%H:%M:%S'),
        'note': '月份由交货日期推导; 数据清洗: 删除空单据编号与业务关闭行; 产线/时效/创建人/单据类型标签来自源表(空则留空); 创建人/单据类型按采购订单编号关联(同PO取最多值); 月产能缺失按默认值补齐',
    }
    DATA = {'meta': meta, 'suppliers': suppliers, 'months': months,
            'capacity': capacity, 'po': po}

    with io.open(OUT, 'w', encoding='utf-8') as f:
        json.dump(DATA, f, ensure_ascii=False, separators=(',', ':'))
    print('RAW:', len(raw), '| dropped empty PO:', dropped_po, '| dropped 业务关闭:', dropped_closed)
    print('KEPT:', len(po), '| suppliers:', len(suppliers), '| months:', months)
    print('产能 real/pending:', real_count, pending_count)
    print('data.json ->', OUT, '(%d bytes)' % os.path.getsize(OUT))

if __name__ == '__main__':
    main()
