import csv
import json
import os
import re
from datetime import datetime, timedelta

TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
SEVEN_DAYS_AGO = TODAY - timedelta(days=7)
THIRTY_DAYS_AGO = TODAY - timedelta(days=30)
THREE_DAYS_AGO = TODAY - timedelta(days=3)

TRACKER_SHEET_URL = (
    'https://docs.google.com/spreadsheets/d/'
    '17-4ObBdkyw9dRi_a4psZJ-p3rm2YnvIZTY922y7mPsY'
    '/export?format=csv&gid=1114317169'
)

def parse_date(raw):
    """Extract date from format like '14:1628/01/2026' (HH:MMDD/MM/YYYY)."""
    if not raw or raw.strip() == '-':
        return None
    raw = raw.strip()
    match = re.search(r'(\d{2}/\d{2}/\d{4})', raw)
    if match:
        try:
            return datetime.strptime(match.group(1), '%d/%m/%Y')
        except ValueError:
            return None
    return None

def parse_euro(raw):
    """Parse euro value like '955.27€', '1,041.22€', '-513.53€'."""
    if not raw or raw.strip() in ('-', '', '#ERROR!'):
        return None
    raw = raw.strip().replace('€', '').replace(',', '')
    try:
        return float(raw)
    except ValueError:
        return None

def load_invoices(path):
    rows = []
    with open(path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        header = next(reader)
        for line in reader:
            if len(line) < 10:
                continue
            received = parse_date(line[0])
            inv_date = parse_date(line[1])
            due_date = parse_date(line[2])
            invoice_num = line[3].strip() if len(line) > 3 else ''
            supplier = line[4].strip() if len(line) > 4 else ''
            inv_total = parse_euro(line[6]) if len(line) > 6 else None
            del_total = parse_euro(line[7]) if len(line) > 7 else None
            diff_raw = line[9].strip() if len(line) > 9 else '-'
            diff_value = parse_euro(diff_raw)
            is_error = diff_raw == '#ERROR!'

            if is_error and inv_total is not None and del_total is not None:
                diff_value = del_total - inv_total

            rows.append({
                'received': received,
                'inv_date': inv_date,
                'due_date': due_date,
                'invoice_num': invoice_num,
                'supplier': supplier,
                'inv_total': inv_total,
                'del_total': del_total,
                'diff_raw': diff_raw,
                'diff_value': diff_value,
            })
    return rows

def compute_metrics(rows):
    total = len(rows)

    overdue = [r for r in rows if r['due_date'] and r['due_date'] < TODAY]
    on_table_7d = [r for r in rows if r['received'] and r['received'] < SEVEN_DAYS_AGO]

    pos_disc = [r for r in rows if r['diff_value'] is not None and r['diff_value'] > 0]
    neg_disc = [r for r in rows if r['diff_value'] is not None and r['diff_value'] < 0]
    on_table_30d = [r for r in rows if r['received'] and r['received'] <= THIRTY_DAYS_AGO]

    def pct(count):
        return round(count / total * 100, 1) if total else 0

    return {
        'total': total,
        'overdue_count': len(overdue),
        'overdue_pct': pct(len(overdue)),
        'over7d_count': len(on_table_7d),
        'over7d_pct': pct(len(on_table_7d)),
        'pos_disc_count': len(pos_disc),
        'pos_disc_pct': pct(len(pos_disc)),
        'neg_disc_count': len(neg_disc),
        'neg_disc_pct': pct(len(neg_disc)),
        'over30d_count': len(on_table_30d),
        'over30d_pct': pct(len(on_table_30d)),
        'overdue_list': overdue,
        'pos_disc_list': pos_disc,
        'neg_disc_list': neg_disc,
        'over30d_list': on_table_30d,
    }

def fmt_date(d):
    return d.strftime('%d/%m/%Y') if d else None

def fmt_date_display(d):
    return d.strftime('%d/%m/%Y') if d else '-'

def serialize_rows(rows):
    import json
    out = []
    for r in rows:
        out.append({
            'received': fmt_date(r['received']),
            'inv_date': fmt_date(r['inv_date']),
            'due_date': fmt_date(r['due_date']),
            'invoice_num': r['invoice_num'],
            'supplier': r['supplier'],
            'inv_total': r['inv_total'],
            'del_total': r['del_total'],
            'diff_value': r['diff_value'],
        })
    return json.dumps(out)

def build_html(m, rows_json, history=None, tracker=None):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Invoice Overview Dashboard</title>
<style>
:root {{
  --blue: #4361ee; --red: #e63946; --orange: #f77f00;
  --teal: #2a9d8f; --coral: #e76f51; --purple: #7209b7; --crimson: #d00000;
  --gray-50: #f9fafb; --gray-100: #f0f2f5; --gray-200: #e5e7eb;
  --gray-400: #9ca3af; --gray-500: #6b7280; --gray-900: #1a1a2e;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: var(--gray-100); color: var(--gray-900); }}

.header {{ background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: #fff; padding: 28px 40px; }}
.header h1 {{ font-size: 1.6rem; font-weight: 600; letter-spacing: -0.02em; }}
.header .subtitle {{ font-size: 0.85rem; color: #a0aec0; margin-top: 4px; }}
.container {{ max-width: 1280px; margin: 0 auto; padding: 24px; }}

/* Toolbar */
.toolbar {{ display: flex; gap: 12px; margin-bottom: 24px; flex-wrap: wrap; align-items: center; }}
.search-box {{ flex: 1; min-width: 220px; padding: 10px 16px; border: 1px solid var(--gray-200); border-radius: 8px; font-size: 0.85rem; outline: none; transition: border-color 0.2s; }}
.search-box:focus {{ border-color: var(--blue); box-shadow: 0 0 0 3px rgba(67,97,238,0.12); }}
.filter-select {{ padding: 10px 16px; border: 1px solid var(--gray-200); border-radius: 8px; font-size: 0.85rem; background: #fff; outline: none; cursor: pointer; min-width: 180px; }}
.filter-select:focus {{ border-color: var(--blue); }}
.reset-btn {{ padding: 10px 20px; border: none; border-radius: 8px; background: var(--gray-200); color: var(--gray-500); font-size: 0.85rem; font-weight: 600; cursor: pointer; transition: all 0.2s; }}
.reset-btn:hover {{ background: #d1d5db; color: var(--gray-900); }}

/* KPI Cards */
.kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 28px; }}
.kpi-card {{ background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); position: relative; overflow: hidden; cursor: pointer; transition: transform 0.15s, box-shadow 0.15s; border: 2px solid transparent; }}
.kpi-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.1); }}
.kpi-card.active {{ border-color: var(--gray-900); }}
.kpi-card::before {{ content: ''; position: absolute; top: 0; left: 0; width: 4px; height: 100%; }}
.kpi-card[data-kpi="total"]::before {{ background: var(--blue); }}
.kpi-card[data-kpi="overdue"]::before {{ background: var(--red); }}
.kpi-card[data-kpi="over7d"]::before {{ background: var(--orange); }}
.kpi-card[data-kpi="pos"]::before {{ background: var(--teal); }}
.kpi-card[data-kpi="neg"]::before {{ background: var(--coral); }}
.kpi-card[data-kpi="over30d"]::before {{ background: var(--crimson); }}
.kpi-card[data-kpi="today"]::before {{ background: #06d6a0; }}
.kpi-card[data-kpi="yesterday"]::before {{ background: #118ab2; }}
.kpi-card[data-kpi="compare"]::before {{ background: #6c63ff; }}
.kpi-card[data-kpi="tracker"]::before {{ background: #ff6b6b; }}
.kpi-card[data-kpi="tracker"] .kpi-value {{ color: #ff6b6b; }}
.tracker-day {{ font-size: 0.72rem; color: var(--gray-500); margin-top: 2px; }}
.kpi-card.month-card::before {{ background: var(--purple); }}
.kpi-label {{ font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--gray-500); font-weight: 600; margin-bottom: 6px; }}
.kpi-value {{ font-size: 1.8rem; font-weight: 700; line-height: 1.1; }}
.kpi-pct {{ font-size: 0.8rem; color: var(--gray-400); margin-top: 4px; }}
.kpi-card[data-kpi="total"] .kpi-value {{ color: var(--blue); }}
.kpi-card[data-kpi="overdue"] .kpi-value {{ color: var(--red); }}
.kpi-card[data-kpi="over7d"] .kpi-value {{ color: var(--orange); }}
.kpi-card[data-kpi="pos"] .kpi-value {{ color: var(--teal); }}
.kpi-card[data-kpi="neg"] .kpi-value {{ color: var(--coral); }}
.kpi-card[data-kpi="over30d"] .kpi-value {{ color: var(--crimson); }}
.kpi-card[data-kpi="compare"] .kpi-value {{ color: #6c63ff; }}
.compare-detail {{ font-size: 0.72rem; color: var(--gray-500); margin-top: 2px; line-height: 1.4; }}
.compare-detail .approved {{ color: #2a9d8f; font-weight: 700; }}
.compare-detail .new-inv {{ color: var(--orange); font-weight: 700; }}
.kpi-card[data-kpi="today"] .kpi-value {{ color: #06d6a0; }}
.kpi-card[data-kpi="yesterday"] .kpi-value {{ color: #118ab2; }}
.kpi-card.month-card .kpi-value {{ color: var(--purple); }}

/* Bar chart */
.section {{ background: #fff; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); margin-bottom: 20px; overflow: hidden; }}
.section-header {{ padding: 16px 24px; border-bottom: 1px solid var(--gray-200); display: flex; align-items: center; gap: 10px; cursor: pointer; user-select: none; transition: background 0.15s; }}
.section-header:hover {{ background: var(--gray-50); }}
.section-header h2 {{ font-size: 0.95rem; font-weight: 600; flex: 1; }}
.section-header .chevron {{ font-size: 0.7rem; color: var(--gray-400); transition: transform 0.25s; }}
.section-header .chevron.collapsed {{ transform: rotate(-90deg); }}
.badge {{ font-size: 0.7rem; padding: 3px 8px; border-radius: 999px; font-weight: 600; color: #fff; white-space: nowrap; }}
.section-body {{ transition: max-height 0.35s ease, opacity 0.25s ease; overflow: hidden; }}
.section-body.collapsed {{ max-height: 0 !important; opacity: 0; }}
.section.highlight {{ animation: flashHighlight 1.2s ease; }}
@keyframes flashHighlight {{ 0%,100% {{ box-shadow: 0 1px 3px rgba(0,0,0,0.08); }} 30% {{ box-shadow: 0 0 0 3px rgba(67,97,238,0.3); }} }}

/* Bar rows */
.bar-row {{ display: flex; align-items: center; gap: 12px; padding: 0 24px 16px; }}
.bar-track {{ flex: 1; height: 10px; background: var(--gray-200); border-radius: 999px; overflow: hidden; }}
.bar-fill {{ height: 100%; border-radius: 999px; transition: width 0.8s ease; }}
.bar-label {{ font-size: 0.75rem; color: var(--gray-500); min-width: 110px; }}

/* Tables */
table {{ width: 100%; border-collapse: collapse; font-size: 0.82rem; }}
th {{ background: var(--gray-50); text-align: left; padding: 10px 16px; font-weight: 600; color: var(--gray-500); text-transform: uppercase; letter-spacing: 0.04em; font-size: 0.72rem; cursor: pointer; user-select: none; white-space: nowrap; transition: background 0.15s; }}
th:hover {{ background: #eef0f3; }}
th .sort-icon {{ display: inline-block; margin-left: 4px; font-size: 0.6rem; opacity: 0.4; }}
th.sort-asc .sort-icon, th.sort-desc .sort-icon {{ opacity: 1; color: var(--blue); }}
td {{ padding: 10px 16px; border-top: 1px solid #f0f0f0; }}
tbody tr {{ transition: background 0.1s; }}
tbody tr:hover td {{ background: #f3f4ff; }}
.urgent {{ color: var(--red); font-weight: 700; }}
.positive {{ color: var(--teal); font-weight: 600; }}
.negative {{ color: var(--coral); font-weight: 600; }}
.no-results {{ padding: 24px; text-align: center; color: var(--gray-400); font-style: italic; }}

@media (max-width: 640px) {{
  .kpi-grid {{ grid-template-columns: 1fr 1fr; }}
  .container {{ padding: 12px; }}
  .header {{ padding: 20px; }}
  .toolbar {{ flex-direction: column; }}
}}
</style>
</head>
<body>

<div class="header">
  <h1>Invoice Overview Dashboard</h1>
  <div class="subtitle">Data as of {TODAY.strftime('%d %B %Y')} &middot; <span id="subtitleCount">{m['total']}</span> invoices on the table</div>
</div>

<div class="container">

  <!-- Toolbar -->
  <div class="toolbar">
    <input type="text" class="search-box" id="searchBox" placeholder="Search invoices, suppliers...">
    <select class="filter-select" id="supplierFilter"><option value="">All Suppliers</option></select>
    <button class="reset-btn" id="resetBtn">Reset Filters</button>
  </div>

  <!-- KPI Cards -->
  <div class="kpi-grid" id="kpiGrid"></div>

  <!-- At a Glance -->
  <div class="section" id="sec-glance">
    <div class="section-header" onclick="toggleSection('glance')">
      <h2>At a Glance</h2>
      <span class="chevron" id="chev-glance">&#9660;</span>
    </div>
    <div class="section-body" id="body-glance"><div style="padding-top:16px" id="barsContainer"></div></div>
  </div>

  <!-- Detail sections rendered by JS -->
  <div id="sectionsContainer"></div>
</div>

<script>
const TODAY = new Date({TODAY.year}, {TODAY.month - 1}, {TODAY.day});
const RAW = {rows_json};
const HISTORY = {json.dumps(history or {{}})};
const TRACKER = {json.dumps(tracker or {{'total':0,'last3d':0,'by_day':{{}},'rows':[]}})};

function getCompareInfo(currentTotal) {{
  const todayKey = TODAY.getFullYear() + '-' + String(TODAY.getMonth()+1).padStart(2,'0') + '-' + String(TODAY.getDate()).padStart(2,'0');
  const dates = Object.keys(HISTORY).filter(d => d < todayKey).sort().reverse();
  if (dates.length === 0) return null;
  const prevDate = dates[0];
  const prevTotal = HISTORY[prevDate];
  const diff = currentTotal - prevTotal;
  const approved = diff < 0 ? Math.abs(diff) : 0;
  const newInv = diff > 0 ? diff : 0;
  const parts = prevDate.split('-');
  const label = parts[2] + '/' + parts[1];
  return {{ prevDate: label, prevTotal, diff, approved, newInv }};
}}

function parseDate(s) {{
  if (!s) return null;
  const p = s.split('/');
  return new Date(+p[2], +p[1]-1, +p[0]);
}}
function daysDiff(a, b) {{ return Math.round((a - b) / 86400000); }}
function fmtEuro(v) {{
  if (v == null) return '-';
  return v.toLocaleString('de-DE', {{minimumFractionDigits:2, maximumFractionDigits:2}}) + '\u20ac';
}}
function fmtDate(s) {{ return s || '-'; }}
function pct(n, t) {{ return t ? Math.round(n/t*1000)/10 : 0; }}

const data = RAW.map(r => ({{
  ...r,
  _received: parseDate(r.received),
  _inv_date: parseDate(r.inv_date),
  _due_date: parseDate(r.due_date),
}}));

const suppliers = [...new Set(data.map(r => r.supplier))].filter(Boolean).sort();
const sel = document.getElementById('supplierFilter');
suppliers.forEach(s => {{ const o = document.createElement('option'); o.value = s; o.textContent = s; sel.appendChild(o); }});

let searchTerm = '';
let supplierVal = '';
let activeKpi = null;
const sortState = {{}};

function filtered() {{
  let rows = data;
  if (searchTerm) {{
    const q = searchTerm.toLowerCase();
    rows = rows.filter(r => (r.invoice_num + ' ' + r.supplier).toLowerCase().includes(q));
  }}
  if (supplierVal) rows = rows.filter(r => r.supplier === supplierVal);
  return rows;
}}

const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const MONTH_COLORS = ['#7209b7','#560bad','#480ca8','#3a0ca3','#3f37c9','#4361ee','#4895ef','#4cc9f0','#72efdd','#2a9d8f','#264653','#e76f51'];

function getMonthKey(d) {{ return d.getFullYear() + '-' + String(d.getMonth()).padStart(2,'0'); }}
function monthLabel(key) {{
  const [y, m] = key.split('-');
  return MONTH_NAMES[+m] + ' ' + y;
}}

const allMonths = [...new Set(data.filter(r => r._inv_date).map(r => getMonthKey(r._inv_date)))].sort();
const currentMonthKey = getMonthKey(TODAY);

const YESTERDAY = new Date(TODAY); YESTERDAY.setDate(YESTERDAY.getDate() - 1);
function sameDay(a, b) {{ return a && b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }}

function computeMetrics(rows) {{
  const total = rows.length;
  const sevenAgo = new Date(TODAY); sevenAgo.setDate(sevenAgo.getDate() - 7);
  const thirtyAgo = new Date(TODAY); thirtyAgo.setDate(thirtyAgo.getDate() - 30);

  const overdue = rows.filter(r => r._due_date && r._due_date < TODAY);
  const over7d = rows.filter(r => r._received && r._received < sevenAgo);
  const pos = rows.filter(r => r.diff_value != null && r.diff_value > 0);
  const neg = rows.filter(r => r.diff_value != null && r.diff_value < 0);
  const over30d = rows.filter(r => r._received && r._received <= thirtyAgo);
  const recToday = rows.filter(r => sameDay(r._received, TODAY));
  const recYesterday = rows.filter(r => sameDay(r._received, YESTERDAY));

  const byMonth = {{}};
  allMonths.forEach(mk => {{
    byMonth[mk] = rows.filter(r => r._inv_date && getMonthKey(r._inv_date) === mk);
  }});

  return {{ total, overdue, over7d, pos, neg, over30d, recToday, recYesterday, byMonth }};
}}

function buildSectionDefs() {{
  const secs = [
    {{ id:'today', title:'Received Today ('+TODAY.toLocaleDateString('en-GB')+')', color:'#06d6a0',
      cols:['Invoice #','Supplier','Invoice Date','Due Date','Amount'],
      keys:['invoice_num','supplier','inv_date','due_date','inv_total'],
      row: r => r,
      fmts: [null,null,r=>fmtDate(r.inv_date),r=>fmtDate(r.due_date),r=>fmtEuro(r.inv_total)],
      cls: [null,null,null,null,null] }},
    {{ id:'yesterday', title:'Received Yesterday ('+YESTERDAY.toLocaleDateString('en-GB')+')', color:'#118ab2',
      cols:['Invoice #','Supplier','Invoice Date','Due Date','Amount'],
      keys:['invoice_num','supplier','inv_date','due_date','inv_total'],
      row: r => r,
      fmts: [null,null,r=>fmtDate(r.inv_date),r=>fmtDate(r.due_date),r=>fmtEuro(r.inv_total)],
      cls: [null,null,null,null,null] }},
    {{ id:'overdue', title:'Overdue Invoices', color:'#e63946',
      cols:['Invoice #','Supplier','Due Date','Days Over','Amount'],
      keys:['invoice_num','supplier','due_date','_daysOver','inv_total'],
      row: r => {{ r._daysOver = r._due_date ? daysDiff(TODAY, r._due_date) : 0; return r; }},
      fmts: [null,null,r=>fmtDate(r.due_date), r=>r._daysOver+'d', r=>fmtEuro(r.inv_total)],
      cls: [null,null,null,'urgent',null] }},
    {{ id:'over7d', title:'On Table > 7 Days', color:'#f77f00',
      cols:['Invoice #','Supplier','Received','Days on Table','Amount'],
      keys:['invoice_num','supplier','received','_daysOn','inv_total'],
      row: r => {{ r._daysOn = r._received ? daysDiff(TODAY, r._received) : 0; return r; }},
      fmts: [null,null,r=>fmtDate(r.received), r=>r._daysOn+'d', r=>fmtEuro(r.inv_total)],
      cls: [null,null,null,'urgent',null] }},
    {{ id:'over30d', title:'On the Table > 30 Days', color:'#d00000',
      cols:['Invoice #','Supplier','Received','Days on Table','Amount'],
      keys:['invoice_num','supplier','received','_daysOn30','inv_total'],
      row: r => {{ r._daysOn30 = r._received ? daysDiff(TODAY, r._received) : 0; return r; }},
      fmts: [null,null,r=>fmtDate(r.received), r=>r._daysOn30+'d', r=>fmtEuro(r.inv_total)],
      cls: [null,null,null,'urgent',null] }},
  ];

  allMonths.forEach((mk, idx) => {{
    const color = MONTH_COLORS[idx % MONTH_COLORS.length];
    secs.push({{
      id: 'month_' + mk.replace('-','_'),
      monthKey: mk,
      title: 'Invoices from ' + monthLabel(mk),
      color: color,
      cols: ['Invoice #','Supplier','Invoice Date','Received','Amount'],
      keys: ['invoice_num','supplier','inv_date','received','inv_total'],
      row: r => r,
      fmts: [null,null,r=>fmtDate(r.inv_date),r=>fmtDate(r.received),r=>fmtEuro(r.inv_total)],
      cls: [null,null,null,null,null],
    }});
  }});

  secs.push(
    {{ id:'pos', title:'Discrepancy (+) \u2014 Delivery > Invoice', color:'#2a9d8f',
      cols:['Invoice #','Supplier','Invoice Total','Delivery Total','Difference'],
      keys:['invoice_num','supplier','inv_total','del_total','diff_value'],
      row: r => r,
      fmts: [null,null,r=>fmtEuro(r.inv_total),r=>fmtEuro(r.del_total),r=>'+'+fmtEuro(Math.abs(r.diff_value))],
      cls: [null,null,null,null,'positive'] }},
    {{ id:'neg', title:'Discrepancy (\u2212) \u2014 Invoice > Delivery', color:'#e76f51',
      cols:['Invoice #','Supplier','Invoice Total','Delivery Total','Difference'],
      keys:['invoice_num','supplier','inv_total','del_total','diff_value'],
      row: r => r,
      fmts: [null,null,r=>fmtEuro(r.inv_total),r=>fmtEuro(r.del_total),r=>fmtEuro(r.diff_value)],
      cls: [null,null,null,null,'negative'] }},
  );

  secs.push(
    {{ id:'tracker', title:'Discrepancy Tracker \u2014 Last 3 Days', color:'#ff6b6b',
      cols:['Date','Supplier','Product'],
      keys:['date','supplier','product'],
      row: r => r,
      fmts: [null,null,null],
      cls: [null,null,null],
      isTracker: true }},
  );

  return secs;
}}

const SECTIONS = buildSectionDefs();

function buildSections() {{
  const c = document.getElementById('sectionsContainer');
  c.innerHTML = '';
  SECTIONS.forEach(s => {{
    c.innerHTML += `
      <div class="section" id="sec-${{s.id}}">
        <div class="section-header" onclick="toggleSection('${{s.id}}')">
          <h2>${{s.title}}</h2>
          <span class="badge" id="badge-${{s.id}}" style="background:${{s.color}}">0</span>
          <span class="chevron" id="chev-${{s.id}}">&#9660;</span>
        </div>
        <div class="section-body" id="body-${{s.id}}">
          <table><thead><tr id="thead-${{s.id}}"></tr></thead><tbody id="tbody-${{s.id}}"></tbody></table>
        </div>
      </div>`;
  }});
  SECTIONS.forEach(s => {{
    const thead = document.getElementById('thead-'+s.id);
    s.cols.forEach((col, i) => {{
      thead.innerHTML += `<th data-sec="${{s.id}}" data-col="${{i}}" onclick="sortTable('${{s.id}}', ${{i}})">${{col}} <span class="sort-icon">\u25B2</span></th>`;
    }});
  }});
}}

function toggleSection(id) {{
  const body = document.getElementById('body-'+id);
  const chev = document.getElementById('chev-'+id);
  body.classList.toggle('collapsed');
  chev.classList.toggle('collapsed');
}}

function sortTable(secId, colIdx) {{
  const key = secId + '-' + colIdx;
  sortState[key] = sortState[key] === 'asc' ? 'desc' : 'asc';
  document.querySelectorAll(`th[data-sec="${{secId}}"]`).forEach(th => th.classList.remove('sort-asc','sort-desc'));
  document.querySelector(`th[data-sec="${{secId}}"][data-col="${{colIdx}}"]`).classList.add('sort-' + sortState[key]);
  render();
}}

function render() {{
  const rows = filtered();
  const m = computeMetrics(rows);

  const cmp = getCompareInfo(m.total);
  const kpis = [
    {{ id:'total', label:'Total Invoices', value:m.total, sub:'on the table', isMonth:false }},
    {{ id:'today', label:'Received Today', value:m.recToday.length, sub: pct(m.recToday.length,m.total)+'% of total', isMonth:false }},
    {{ id:'yesterday', label:'Received Yesterday', value:m.recYesterday.length, sub: pct(m.recYesterday.length,m.total)+'% of total', isMonth:false }},
    {{ id:'overdue', label:'Overdue', value:m.overdue.length, sub: pct(m.overdue.length,m.total)+'% of total', isMonth:false }},
    {{ id:'over7d', label:'On Table > 7 Days', value:m.over7d.length, sub: pct(m.over7d.length,m.total)+'% of total', isMonth:false }},
    {{ id:'pos', label:'Discrepancy (+)', value:m.pos.length, sub: pct(m.pos.length,m.total)+'% of total', isMonth:false }},
    {{ id:'neg', label:'Discrepancy (\u2212)', value:m.neg.length, sub: pct(m.neg.length,m.total)+'% of total', isMonth:false }},
    {{ id:'over30d', label:'On Table > 30 Days', value:m.over30d.length, sub: pct(m.over30d.length,m.total)+'% of total', isMonth:false }},
  ];
  allMonths.forEach(mk => {{
    const secId = 'month_' + mk.replace('-','_');
    const cnt = (m.byMonth[mk] || []).length;
    kpis.push({{ id:secId, label:monthLabel(mk), value:cnt, sub: pct(cnt,m.total)+'% of total', isMonth:true }});
  }});
  const grid = document.getElementById('kpiGrid');
  grid.innerHTML = '';
  kpis.forEach(k => {{
    const active = activeKpi === k.id ? ' active' : '';
    const monthCls = k.isMonth ? ' month-card' : '';
    grid.innerHTML += `<div class="kpi-card${{active}}${{monthCls}}" data-kpi="${{k.id}}" onclick="kpiClick('${{k.id}}')">
      <div class="kpi-label">${{k.label}}</div><div class="kpi-value">${{k.value}}</div><div class="kpi-pct">${{k.sub}}</div></div>`;
  }});
  if (cmp) {{
    const arrow = cmp.diff < 0 ? '\u2193' : cmp.diff > 0 ? '\u2191' : '\u2194';
    const diffAbs = Math.abs(cmp.diff);
    let detail = `<div class="compare-detail">Previous (${{cmp.prevDate}}): <b>${{cmp.prevTotal}}</b>`;
    if (cmp.approved > 0) detail += `<br><span class="approved">${{cmp.approved}} approved / removed</span>`;
    if (cmp.newInv > 0) detail += `<br><span class="new-inv">${{cmp.newInv}} new invoices added</span>`;
    if (cmp.diff === 0) detail += `<br>No change since last update`;
    detail += `</div>`;
    grid.innerHTML += `<div class="kpi-card" data-kpi="compare" style="grid-column: span 2;">
      <div class="kpi-label">Day-over-Day Change</div>
      <div class="kpi-value">${{arrow}} ${{diffAbs}}</div>
      <div class="kpi-pct">${{m.total}} today vs ${{cmp.prevTotal}} previous</div>${{detail}}</div>`;
  }}

  if (TRACKER.last3d > 0 || TRACKER.total > 0) {{
    let dayBreakdown = Object.entries(TRACKER.by_day).sort().reverse()
      .map(([d,c]) => `<span class="tracker-day">${{d}}: <b>${{c}}</b></span>`).join(' &middot; ');
    grid.innerHTML += `<div class="kpi-card" data-kpi="tracker" onclick="kpiClick('tracker')" style="grid-column: span 2;">
      <div class="kpi-label">Discrepancy Tracker (Last 3 Days)</div>
      <div class="kpi-value">${{TRACKER.last3d}}</div>
      <div class="kpi-pct">${{TRACKER.total}} total tracked discrepancies</div>
      <div style="margin-top:4px">${{dayBreakdown}}</div></div>`;
  }}

  document.getElementById('subtitleCount').textContent = m.total;

  const bars = [
    ['Received today', m.recToday.length, m.total, '#06d6a0'],
    ['Received yesterday', m.recYesterday.length, m.total, '#118ab2'],
    ['Overdue', m.overdue.length, m.total, '#e63946'],
    ['> 7 days on table', m.over7d.length, m.total, '#f77f00'],
    ['Discrepancy (+)', m.pos.length, m.total, '#2a9d8f'],
    ['Discrepancy (\u2212)', m.neg.length, m.total, '#e76f51'],
    ['> 30 days on table', m.over30d.length, m.total, '#d00000'],
  ];
  allMonths.forEach((mk, idx) => {{
    const cnt = (m.byMonth[mk] || []).length;
    bars.push([monthLabel(mk), cnt, m.total, MONTH_COLORS[idx % MONTH_COLORS.length]]);
  }});
  const bc = document.getElementById('barsContainer');
  bc.innerHTML = '';
  bars.forEach(([label, count, total, color]) => {{
    const p = pct(count, total);
    bc.innerHTML += `<div class="bar-row"><div class="bar-label">${{label}}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${{p}}%;background:${{color}}"></div></div>
      <div class="bar-label" style="text-align:right">${{count}} pcs (${{p}}%)</div></div>`;
  }});

  const sectionData = {{
    today: m.recToday, yesterday: m.recYesterday,
    overdue: m.overdue, over7d: m.over7d, pos: m.pos, neg: m.neg,
    over30d: m.over30d,
    tracker: TRACKER.rows || []
  }};
  allMonths.forEach(mk => {{
    sectionData['month_' + mk.replace('-','_')] = m.byMonth[mk] || [];
  }});

  SECTIONS.forEach(sec => {{
    let sRows = (sectionData[sec.id] || []).map(sec.row);
    const sk = sec.id + '-' + Object.keys(sortState).filter(k=>k.startsWith(sec.id+'-')).find(k=>sortState[k]) ;
    for (const k in sortState) {{
      if (!k.startsWith(sec.id+'-')) continue;
      const ci = +k.split('-').pop();
      const key = sec.keys[ci];
      const dir = sortState[k] === 'desc' ? -1 : 1;
      sRows.sort((a,b) => {{
        let va = a[key], vb = b[key];
        if (va == null) va = '';
        if (vb == null) vb = '';
        if (typeof va === 'number' && typeof vb === 'number') return (va - vb)*dir;
        return String(va).localeCompare(String(vb))*dir;
      }});
    }}

    document.getElementById('badge-'+sec.id).textContent = sRows.length + ' pcs';
    const tbody = document.getElementById('tbody-'+sec.id);
    if (!sRows.length) {{
      tbody.innerHTML = '<tr><td colspan="'+sec.cols.length+'" class="no-results">No invoices match the current filters</td></tr>';
      return;
    }}
    tbody.innerHTML = '';
    sRows.forEach(r => {{
      let tr = '<tr>';
      sec.cols.forEach((_, i) => {{
        const val = sec.fmts[i] ? sec.fmts[i](r) : (r[sec.keys[i]] ?? '-');
        const c = sec.cls[i] ? ` class="${{sec.cls[i]}}"` : '';
        tr += `<td${{c}}>${{val}}</td>`;
      }});
      tr += '</tr>';
      tbody.innerHTML += tr;
    }});
  }});
}}

function kpiClick(id) {{
  if (activeKpi === id) {{ activeKpi = null; }}
  else {{ activeKpi = id; }}
  render();
  if (activeKpi && activeKpi !== 'total') {{
    const sec = document.getElementById('sec-' + id);
    if (sec) {{
      const body = document.getElementById('body-' + id);
      if (body.classList.contains('collapsed')) toggleSection(id);
      sec.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
      sec.classList.remove('highlight');
      void sec.offsetWidth;
      sec.classList.add('highlight');
    }}
  }}
}}

document.getElementById('searchBox').addEventListener('input', e => {{ searchTerm = e.target.value; render(); }});
document.getElementById('supplierFilter').addEventListener('change', e => {{ supplierVal = e.target.value; render(); }});
document.getElementById('resetBtn').addEventListener('click', () => {{
  searchTerm = ''; supplierVal = ''; activeKpi = null;
  document.getElementById('searchBox').value = '';
  document.getElementById('supplierFilter').value = '';
  for (const k in sortState) delete sortState[k];
  document.querySelectorAll('th').forEach(th => th.classList.remove('sort-asc','sort-desc'));
  render();
}});

buildSections();
render();
</script>
</body>
</html>"""

def load_tracker_data():
    """Load discrepancy tracker from local CSV or fetch from Google Sheets."""
    script_dir = os.path.dirname(__file__)
    local_path = os.path.join(script_dir, '🇵🇹 Invoice Discrepancy Tracker [PT-2026] - 🔎 Tracker - Discrepancies.csv')

    if not os.path.exists(local_path):
        try:
            import urllib.request
            urllib.request.urlretrieve(TRACKER_SHEET_URL, local_path)
        except Exception as e:
            print(f"Could not fetch tracker sheet: {e}")
            return {'total': 0, 'last3d': 0, 'by_day': {}, 'rows': []}

    total = 0
    last3d_rows = []
    by_day = {}
    with open(local_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        all_rows = list(reader)

    for row in all_rows[4:]:
        if len(row) < 3 or not row[2].strip():
            continue
        try:
            d = datetime.strptime(row[2].strip(), '%d-%b-%Y')
        except ValueError:
            continue
        total += 1
        if d >= THREE_DAYS_AGO:
            supplier = row[0].strip() if row[0].strip() else '-'
            product = row[7].strip() if len(row) > 7 and row[7].strip() else '-'
            day_str = d.strftime('%d/%m/%Y')
            by_day[day_str] = by_day.get(day_str, 0) + 1
            last3d_rows.append({
                'date': day_str,
                'supplier': supplier,
                'product': product,
            })

    return {
        'total': total,
        'last3d': len(last3d_rows),
        'by_day': by_day,
        'rows': last3d_rows,
    }


def load_history():
    import json, os
    path = os.path.join(os.path.dirname(__file__), 'invoice_history.json')
    if os.path.exists(path):
        try:
            with open(path) as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


if __name__ == '__main__':
    import os
    csv_path = os.path.join(os.path.dirname(__file__), 'invoices overview - Sheet1.csv')
    rows = load_invoices(csv_path)
    metrics = compute_metrics(rows)
    rows_json = serialize_rows(rows)
    history = load_history()
    tracker = load_tracker_data()
    print(f"Tracker: {tracker['total']} total, {tracker['last3d']} in last 3 days")
    html = build_html(metrics, rows_json, history, tracker)
    out_path = os.path.join(os.path.dirname(__file__), 'dashboard.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Dashboard generated: {out_path}")
    print(f"Total invoices: {metrics['total']}")
    print(f"Overdue: {metrics['overdue_count']} ({metrics['overdue_pct']}%)")
    print(f"On table > 7 days: {metrics['over7d_count']} ({metrics['over7d_pct']}%)")
    print(f"Discrepancy (+): {metrics['pos_disc_count']} ({metrics['pos_disc_pct']}%)")
    print(f"Discrepancy (-): {metrics['neg_disc_count']} ({metrics['neg_disc_pct']}%)")
    print(f"On table > 30 days: {metrics['over30d_count']} ({metrics['over30d_pct']}%)")
