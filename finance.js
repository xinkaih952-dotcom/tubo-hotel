/* =================================================================
 * 涂泊酒店（南宁东站店）· 收支利润（店长页面）
 * 依赖：data.js（历史营收）、finance-data.js（历史成本）、common.js（配置与共用代码）
 * 功能：月度成本录入、利润与毛利率看板、成本结构饼图、月度利润趋势、
 *       月度利润表（与酒店利润表格式一致）、导入前台 Excel、导出
 * ================================================================= */

const $ = id => document.getElementById(id);
const HISTORY_COST_DATA = (typeof HISTORY_COSTS !== 'undefined' && Array.isArray(HISTORY_COSTS)) ? HISTORY_COSTS : [];
const HISTORY_COST_MAP = Object.fromEntries(HISTORY_COST_DATA.map(c => [c.month, c]));

const state = {
  month: null,        // 当前查看的月份
  pieScope: 'month'   // 成本结构：本月 / 本年累计
};

/* -----------------------------------------------------------------
 * 一、月度成本存储（手动录入优先，其次用利润表历史数据）
 * ----------------------------------------------------------------- */
const CostStore = (() => {
  let data = { version: 1, months: {} };
  function clean(obj) {
    const c = {};
    COST_KEYS.forEach(k => { c[k] = Number(obj && obj[k]) || 0; });
    return c;
  }
  return {
    load() {
      const d = Utils.readJSON(CONFIG.COST_KEY);
      data = d && d.months ? d : { version: 1, months: {} };
    },
    /** 取某月成本；没有返回 null */
    get(month) {
      if (data.months[month]) return clean(data.months[month]);
      if (HISTORY_COST_MAP[month]) return clean(HISTORY_COST_MAP[month]);
      return null;
    },
    source(month) {
      if (data.months[month]) return HISTORY_COST_MAP[month] ? 'edited' : 'manual';
      return HISTORY_COST_MAP[month] ? 'history' : null;
    },
    save(month, values) {
      data.months[month] = Object.assign(clean(values), { updatedAt: new Date().toISOString() });
      return Utils.writeJSON(CONFIG.COST_KEY, data);
    },
    remove(month) {
      delete data.months[month];
      return Utils.writeJSON(CONFIG.COST_KEY, data);
    },
    months: () => [...new Set([...Object.keys(data.months), ...Object.keys(HISTORY_COST_MAP)])].sort(),
    all: () => data
  };
})();

/* -----------------------------------------------------------------
 * 二、导入的前台 Excel 数据
 * ----------------------------------------------------------------- */
const ImportStore = (() => {
  let data = { version: 1, records: [], files: [] };
  return {
    load() {
      const d = Utils.readJSON(CONFIG.IMPORT_KEY);
      data = d && Array.isArray(d.records) ? d : { version: 1, records: [], files: [] };
      LedgerStore.setExtra(data.records);
    },
    add(records, fileName) {
      const byDate = Object.fromEntries(data.records.map(r => [r.date, r]));
      records.forEach(r => { byDate[r.date] = r; });   // 同一天以新导入的为准
      data.records = Object.values(byDate);
      data.files.push({ name: fileName, days: records.length, at: new Date().toISOString() });
      Utils.writeJSON(CONFIG.IMPORT_KEY, data);
      LedgerStore.setExtra(data.records);
    },
    clear() {
      data = { version: 1, records: [], files: [] };
      Utils.writeJSON(CONFIG.IMPORT_KEY, data);
      LedgerStore.setExtra([]);
    },
    info: () => data
  };
})();

/** 解析前台导出的 Excel（每日营收表），返回每日记录数组 */
function parseLedgerWorkbook(wb) {
  const sheetName = wb.SheetNames.includes('每日营收') ? '每日营收' : wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: '' });
  const headerIdx = rows.findIndex(r => r.includes('日期') && r.includes('出租间数'));
  if (headerIdx < 0) throw new Error('没有找到「日期」「出租间数」表头，请确认是前台营收台账导出的文件');
  const header = rows[headerIdx].map(h => String(h).trim());
  const labelToKey = { '出租间数': 'roomSold', '平均房价': 'adr', '备注': 'remark', '明细表': 'sysDetail' };
  [...ROOM_CHANNELS, ...NON_ROOM_ITEMS].forEach(c => {
    labelToKey[c.label] = c.key;
    (c.legacy || []).forEach(l => { labelToKey[l] = c.key; });   // 旧版本导出的列名
  });
  const dateCol = header.indexOf('日期');

  const out = [];
  rows.slice(headerIdx + 1).forEach(row => {
    let d = row[dateCol];
    if (typeof d === 'number') {            // Excel 日期序号（文件被 Excel 另存过）
      const p = XLSX.SSF.parse_date_code(d);
      d = p ? `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}` : '';
    }
    d = String(d).trim().replace(/\//g, '-');
    const m = d.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!m) return;                         // 跳过合计行等
    const rec = { date: `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` };
    header.forEach((h, i) => {
      const key = labelToKey[h];
      if (!key) return;
      if (key === 'remark') rec[key] = String(row[i] || '');
      else if (key === 'sysDetail') rec[key] = row[i] === '' ? null : Number(row[i]);
      else rec[key] = Number(row[i]) || 0;
    });
    out.push(RecordModel.normalize(rec));
  });
  return out;
}

async function handleImport(e) {
  const files = [...e.target.files];
  if (!files.length) return;
  if (typeof XLSX === 'undefined') { showToast('Excel 组件没有加载成功，请检查网络后刷新页面再试', true); return; }
  let total = 0;
  for (const f of files) {
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
      const recs = parseLedgerWorkbook(wb);
      if (!recs.length) { showToast(`${f.name} 里没有每日记录`, true); continue; }
      ImportStore.add(recs, f.name);
      total += recs.length;
    } catch (err) {
      showToast(`${f.name} 导入失败：${err.message}`, true);
    }
  }
  e.target.value = '';
  if (total) {
    showToast(`已导入 ${total} 天的营收记录`);
    renderAll();
  }
}

/* -----------------------------------------------------------------
 * 二·五、月底渠道结算对账
 *   系统录入合计：每日记录（本机 + 导入的前台 Excel）；没有时用营收表历史每日数据
 *   实际结算：手动录入优先，其次用营收表「实际结算总额」
 *   差异 = 实际结算 − 系统录入（正数 = 实际到账比系统录入多，通常是漏录）
 * ----------------------------------------------------------------- */
const SETTLE_MAP = Object.fromEntries(HISTORY_SETTLE_DATA.map(x => [x.month, x]));

const SettleStore = (() => {
  let data = { version: 1, months: {} };
  const persist = () => Utils.writeJSON(CONFIG.SETTLE_KEY, data);
  return {
    load() { const d = Utils.readJSON(CONFIG.SETTLE_KEY); data = d && d.months ? d : { version: 1, months: {} }; },
    manual: month => data.months[month] || null,
    /** 取某月结算：{ values: {key: 金额}, note, source } 或 null */
    get(month) {
      const m = data.months[month];
      if (m) return { values: m.values || {}, note: m.note || '', source: SETTLE_MAP[month] ? 'edited' : 'manual' };
      const h = SETTLE_MAP[month];
      if (h) {
        const values = {};
        DETAIL_KEYS.forEach(k => { if (h[k] !== undefined) values[k] = h[k]; });
        return { values, note: h.note || '', source: 'history' };
      }
      return null;
    },
    save(month, values, note) { data.months[month] = { values, note, updatedAt: new Date().toISOString() }; return persist(); },
    remove(month) { delete data.months[month]; return persist(); },
    months: () => [...new Set([...Object.keys(data.months), ...Object.keys(SETTLE_MAP)])].sort()
  };
})();

/** 某月各渠道「系统录入合计」 */
function systemTotals(month) {
  const ledger = LedgerStore.byMonth(month);
  const src = ledger.length ? ledger : DailyHistory.byMonth(month);
  if (!src.length) return null;
  const t = { days: src.length, source: ledger.length ? 'ledger' : 'history' };
  DETAIL_KEYS.forEach(k => { t[k] = Utils.round2(src.reduce((s, r) => s + (r[k] || 0), 0)); });
  return t;
}

/** 利润表用的营收：营收表历史月份用利润表；之后的月份，填了实际结算的渠道用结算数 */
function revenueForPL(month) {
  const rev = Monthly.get(month);
  if (!rev || rev.source === 'history') return rev;
  const st = SettleStore.manual(month);
  if (!st || !st.values || !Object.keys(st.values).length) return rev;
  const m = Object.assign({}, rev);
  Object.entries(st.values).forEach(([k, v]) => { m[k] = Number(v) || 0; });
  const out = Monthly.finish(m, rev.adr);
  out.settled = true;
  return out;
}

const settleInputs = {};

function renderSettle() {
  const month = state.month;
  $('settleTitle').textContent = `渠道结算对账 · ${Utils.monthCn(month)}`;
  const sys = systemTotals(month);
  const st = SettleStore.get(month);
  const used = DETAIL_KEYS.filter(k => REV_ITEM[k].active !== false || (sys && sys[k]) || (st && st.values[k]));
  $('settleSource').textContent = !st ? '还没有填实际结算' : (st.source === 'history' ? '实际结算来自营收表' : st.source === 'edited' ? '已修改（原数据来自营收表）' : '已录入');
  $('btnSettleRestore').hidden = !(st && st.source === 'edited');

  $('settleTable').innerHTML = `
    <thead><tr><th>渠道</th><th class="r">系统录入合计</th><th class="r" style="width:200px">实际结算</th><th class="r">差异（结算 − 系统）</th></tr></thead>
    <tbody>${used.map(k => `
      <tr><td>${REV_ITEM[k].label}</td>
        <td class="r money-cell">${sys ? '¥' + Utils.money(sys[k]) : '—'}</td>
        <td class="r"><div class="money"><input class="input" type="number" step="0.01" inputmode="decimal" id="st_${k}" data-key="${k}" placeholder="未填"
          value="${st && st.values[k] !== undefined ? st.values[k] : ''}" style="text-align:right"></div></td>
        <td class="r" id="stdiff_${k}"></td></tr>`).join('')}
      <tr style="font-weight:600;background:var(--surface-2)"><td>合计</td><td class="r money-cell" id="stSysTotal"></td><td class="r money-cell" id="stActTotal"></td><td class="r" id="stDiffTotal"></td></tr>
    </tbody>`;
  used.forEach(k => { settleInputs[k] = $('st_' + k); settleInputs[k].addEventListener('input', updateSettleDiffs); });
  Object.keys(settleInputs).forEach(k => { if (!used.includes(k)) delete settleInputs[k]; });
  $('settleNote').value = st ? st.note : '';
  $('settleSysInfo').textContent = !sys ? '这个月还没有每日营收数据，无法算出系统录入合计。'
    : `系统录入合计来自${sys.source === 'ledger' ? '前台每日记录' : '营收表每日数据'}，共 ${sys.days} 天。`
      + (Monthly.historyMap[month] ? '' : '保存实际结算后，利润表里这些渠道的收入会改用结算数。');
  updateSettleDiffs();
  renderSettleHistory();
}

function updateSettleDiffs() {
  const sys = systemTotals(state.month);
  let sumSys = 0, sumAct = 0, any = false;
  Object.entries(settleInputs).forEach(([k, input]) => {
    const raw = input.value.trim();
    const sv = sys ? sys[k] : 0;
    sumSys += sv;
    const cell = $('stdiff_' + k);
    if (raw === '' || !Number.isFinite(Number(raw))) { cell.innerHTML = '<span class="src">—</span>'; sumAct += sv; return; }
    any = true;
    const act = Number(raw);
    sumAct += act;
    cell.innerHTML = diffCell(Utils.round2(act - sv));
  });
  $('stSysTotal').textContent = '¥' + Utils.money(sumSys);
  $('stActTotal').textContent = any ? '¥' + Utils.money(sumAct) : '—';
  $('stDiffTotal').innerHTML = any ? diffCell(Utils.round2(sumAct - sumSys)) : '';
}

function diffCell(d) {
  if (Math.abs(d) < 0.005) return '<span class="recon ok">✓ 0</span>';
  return `<span class="recon gap">${d > 0 ? '+' : '−'}¥${Utils.money(Math.abs(d))}</span>`;
}

function handleSettleSave(e) {
  e.preventDefault();
  const values = {};
  let ok = true;
  Object.entries(settleInputs).forEach(([k, input]) => {
    if (input.value.trim() === '') return;
    const res = Utils.readAmount(input);
    if (!res.ok) { ok = false; input.classList.add('has-error'); } else { input.classList.remove('has-error'); values[k] = res.value; }
  });
  if (!ok) { showToast('有金额格式不对，请检查标红的格子', true); return; }
  SettleStore.save(state.month, values, $('settleNote').value.trim());
  renderAll();
  showToast(`已保存 ${Utils.monthCn(state.month)} 实际结算`);
}

/** 历月结算差异一览 */
function renderSettleHistory() {
  const months = SettleStore.months().filter(m => systemTotals(m)).reverse();
  if (!months.length) { $('settleHistory').innerHTML = ''; $('settleInsight').textContent = ''; return; }
  const rows = months.map(m => {
    const sys = systemTotals(m), st = SettleStore.get(m);
    const d = {};
    DETAIL_KEYS.forEach(k => { d[k] = st.values[k] === undefined ? null : Utils.round2(st.values[k] - sys[k]); });
    return { m, d, note: st.note };
  });
  const cols = DETAIL_KEYS.filter(k => rows.some(r => r.d[k] !== null && Math.abs(r.d[k]) >= 0.005));
  const fmt = v => v === null ? '<span class="src">—</span>' : (Math.abs(v) < 0.005 ? '<span class="src">0</span>'
    : `<span style="color:var(--warn);font-weight:600">${v > 0 ? '+' : '−'}${Utils.money(Math.abs(v))}</span>`);
  $('settleHistory').innerHTML = `
    <thead><tr><th>月份</th>${cols.map(k => `<th class="r">${REV_ITEM[k].label}</th>`).join('')}<th>差异说明</th></tr></thead>
    <tbody>${rows.map(r => `<tr data-month="${r.m}" style="cursor:pointer"><td class="date-cell">${Utils.monthCn(r.m)}</td>
      ${cols.map(k => `<td class="r money-cell">${fmt(r.d[k])}</td>`).join('')}
      <td class="remark-cell" style="max-width:320px" title="${Utils.escape(r.note)}">${r.note ? Utils.escape(r.note) : '—'}</td></tr>`).join('')}
      <tr style="font-weight:600;background:var(--surface-2)"><td>累计</td>${cols.map(k => `<td class="r money-cell">${fmt(Utils.round2(rows.reduce((t, r) => t + (r.d[k] || 0), 0)))}</td>`).join('')}<td></td></tr>
    </tbody>`;
  // 一句话提示：哪个渠道持续有差额
  const facts = cols.map(k => {
    const vals = rows.map(r => r.d[k]).filter(v => v !== null);
    const pos = vals.filter(v => v > 0.005).length;
    const total = vals.reduce((t, v) => t + v, 0);
    return { k, pos, n: vals.length, total };
  }).filter(f => f.n >= 3 && f.pos / f.n >= 0.8).sort((a, b) => b.total - a.total);
  $('settleInsight').textContent = facts.length
    ? facts.map(f => `${REV_ITEM[f.k].label}：${f.n} 个月里有 ${f.pos} 个月实际到账比系统录入多，累计多 ¥${Utils.money(f.total)}`).join('；') + '。说明这一项经常漏录，建议前台每天核对。'
    : '';
}

/* -----------------------------------------------------------------
 * 三、利润计算（与酒店利润表结构一致）
 * ----------------------------------------------------------------- */
/** 某组成本合计（抵减项相减） */
function groupSum(cost, groupId) {
  const g = COST_GROUPS.find(x => x.id === groupId);
  return Utils.round2(g.items.reduce((s, i) => s + (i.negative ? -1 : 1) * (cost[i.key] || 0), 0));
}

/** 计算某月损益；营收或成本缺失时返回 null */
function computePL(month) {
  const rev = revenueForPL(month);
  const cost = CostStore.get(month);
  if (!rev || !cost) return null;
  const g = id => groupSum(cost, id);
  const R1 = rev.revenueRoom, R2 = rev.revenueNonRoom, R = rev.revenueTotal;
  const mainCost = Utils.round2(g('ctrl') + g('nonctrl'));
  const expense = Utils.round2(g('sales') + g('admin') + g('fin') + g('tax'));
  const nonMainCost = g('nonmain');
  const below = g('below');
  const mainProfit = Utils.round2(R1 - mainCost - expense);
  const nonMainProfit = Utils.round2(R2 - nonMainCost);
  const opProfit = Utils.round2(mainProfit + nonMainProfit);
  const net = Utils.round2(opProfit - below);
  const totalCost = Utils.round2(mainCost + expense + nonMainCost + below);
  const gross = Utils.round2(R - mainCost - nonMainCost);
  // 成本结构：直接按利润表的分组（可控费用、非可控费用、销售费用……）
  const cats = {};
  COST_GROUPS.forEach(x => { cats[x.id] = Math.max(0, g(x.id)); });
  return {
    month, rev, cost, R1, R2, R, mainCost, expense, nonMainCost, below,
    groups: Object.fromEntries(COST_GROUPS.map(x => [x.id, g(x.id)])),
    mainProfit, nonMainProfit, opProfit, net, totalCost, gross,
    grossMargin: R > 0 ? gross / R : null,
    netMargin: R > 0 ? net / R : null,
    cats, partial: rev.partial, costSource: CostStore.source(month)
  };
}

/** 所有可看的月份（有营收或有成本） */
function allMonths() {
  return [...new Set([...Monthly.allMonths(), ...CostStore.months()])].sort();
}
function plMonths() { return allMonths().filter(m => computePL(m)); }

/* -----------------------------------------------------------------
 * 四、看板
 * ----------------------------------------------------------------- */
function renderMonthSelect() {
  const months = allMonths().slice().reverse();
  if (!state.month || !months.includes(state.month)) {
    const pl = plMonths();
    state.month = pl.length ? pl[pl.length - 1] : (months[0] || Utils.thisMonth());
  }
  $('monthSelect').innerHTML = months.map(m => {
    const tag = computePL(m) ? '' : (CostStore.get(m) ? '（缺营收）' : '（缺成本）');
    return `<option value="${m}" ${m === state.month ? 'selected' : ''}>${Utils.monthCn(m)}${tag}</option>`;
  }).join('');
}

function renderKpis() {
  const m = state.month;
  const pl = computePL(m);
  const prev = computePL(Utils.prevYearMonth(m));
  const rev = Monthly.get(m);
  const cost = CostStore.get(m);
  const banner = $('partialBanner');

  if (!pl) {
    banner.hidden = false;
    banner.textContent = !rev && !cost ? `${Utils.monthCn(m)}还没有营收和成本数据。`
      : (!cost ? `${Utils.monthCn(m)}还没有录入成本，请在下方「月度成本录入」填写后查看利润。`
               : `${Utils.monthCn(m)}还没有营收数据：请导入前台导出的当月 Excel。`);
    $('kRevenue').innerHTML = rev ? `<span class="cur">¥</span>${Utils.money(rev.revenueTotal)}` : '—';
    $('kRevenueFoot').textContent = rev ? `房费 ¥${Utils.money(rev.revenueRoom)}` : '—';
    ['kCost', 'kNet', 'kGross', 'kNetMargin'].forEach(id => { $(id).textContent = '—'; $(id).classList.remove('neg'); });
    ['kCostFoot', 'kNetFoot', 'kGrossFoot', 'kNetMarginFoot'].forEach(id => { $(id).textContent = ''; });
    return;
  }

  banner.hidden = !pl.partial;
  if (pl.partial) banner.textContent = `${Utils.monthCn(m)}营收只录入了 ${pl.rev.days}/${pl.rev.fullDays} 天，成本是整月的，所以利润偏低，月底录满后再看。`;

  const vsPrev = (now, p, type, invert) => p === null || p === undefined
    ? '<span>去年同月无数据</span>'
    : `<span>去年 ${type === 'pp' ? Utils.percent(p) : '¥' + Utils.money(p)}</span>${deltaChip(now, p, type, invert)}`;

  $('kRevenue').innerHTML = `<span class="cur">¥</span>${Utils.money(pl.R)}`;
  $('kRevenueFoot').innerHTML = vsPrev(pl.R, prev && prev.R, 'pct');
  $('kCost').innerHTML = `<span class="cur">¥</span>${Utils.money(pl.totalCost)}`;
  $('kCostFoot').innerHTML = vsPrev(pl.totalCost, prev && prev.totalCost, 'pct', true);
  $('kNet').innerHTML = `${pl.net < 0 ? '−' : ''}<span class="cur">¥</span>${Utils.money(Math.abs(pl.net))}`;
  $('kNet').classList.toggle('neg', pl.net < 0);
  $('kNetFoot').innerHTML = prev ? `<span>去年 ${prev.net < 0 ? '−' : ''}¥${Utils.money(Math.abs(prev.net))}</span>` : '<span>去年同月无数据</span>';
  $('kGross').innerHTML = pl.grossMargin === null ? '—' : `${(pl.grossMargin * 100).toFixed(1)}<span class="unit">%</span>`;
  $('kGrossFoot').innerHTML = vsPrev(pl.grossMargin, prev && prev.grossMargin, 'pp');
  $('kNetMargin').innerHTML = pl.netMargin === null ? '—' : `${(pl.netMargin * 100).toFixed(1)}<span class="unit">%</span>`;
  $('kNetMargin').classList.toggle('neg', pl.netMargin !== null && pl.netMargin < 0);
  $('kNetMarginFoot').innerHTML = vsPrev(pl.netMargin, prev && prev.netMargin, 'pp');
}

/* -----------------------------------------------------------------
 * 五、成本结构饼图（环形）
 * ----------------------------------------------------------------- */
function pieData() {
  const m = state.month;
  let months = [m];
  if (state.pieScope === 'ytd') months = plMonths().filter(x => x.slice(0, 4) === m.slice(0, 4) && x <= m && (x === m || !computePL(x).partial));
  const pls = months.map(computePL).filter(Boolean);
  const cats = {};
  COST_GROUPS.forEach(c => { cats[c.id] = pls.reduce((s, p) => s + p.cats[c.id], 0); });
  return {
    label: state.pieScope === 'ytd' ? `${m.slice(0, 4)}年 1–${Number(m.slice(5))}月` : Utils.monthCn(m),
    months: pls.length,
    revenue: pls.reduce((s, p) => s + p.R, 0),
    items: COST_GROUPS.map(c => ({ id: c.id, label: c.short, color: c.color, value: Utils.round2(cats[c.id]) }))
  };
}

function renderPie() {
  const d = pieData();
  const svg = $('pieChart');
  const total = d.items.reduce((s, i) => s + i.value, 0);
  const S = 260, cx = S / 2, cy = S / 2, R = 118, r = 74;
  svg.setAttribute('viewBox', `0 0 ${S} ${S}`);

  if (total <= 0) {
    svg.innerHTML = `<circle cx="${cx}" cy="${cy}" r="${(R + r) / 2}" fill="none" stroke="var(--line)" stroke-width="${R - r}"/>
      <text class="center-k" x="${cx}" y="${cy + 5}" text-anchor="middle">暂无成本数据</text>`;
    $('pieLegend').innerHTML = '';
    return;
  }

  // 从 12 点方向顺时针，按利润表分组的固定顺序（颜色跟随分组）
  const pt = (a, rad) => [cx + rad * Math.sin(a), cy - rad * Math.cos(a)];
  let a0 = 0, out = '';
  const shown = d.items.filter(i => i.value > 0);
  shown.forEach(i => {
    const a1 = a0 + (i.value / total) * Math.PI * 2;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const [x0, y0] = pt(a0, R), [x1, y1] = pt(a1, R), [x2, y2] = pt(a1, r), [x3, y3] = pt(a0, r);
    const path = shown.length === 1
      ? `M${cx},${cy - R}A${R},${R} 0 1 1 ${cx - 0.01},${cy - R}L${cx - 0.01},${cy - r}A${r},${r} 0 1 0 ${cx},${cy - r}Z`
      : `M${x0},${y0}A${R},${R} 0 ${large} 1 ${x1},${y1}L${x2},${y2}A${r},${r} 0 ${large} 0 ${x3},${y3}Z`;
    out += `<path class="slice" data-id="${i.id}" d="${path}" fill="${i.color}" stroke="var(--surface)" stroke-width="2"/>`;
    a0 = a1;
  });
  out += `<text class="center-k" x="${cx}" y="${cy - 8}" text-anchor="middle">成本费用合计</text>
    <text class="center-v" x="${cx}" y="${cy + 18}" text-anchor="middle">¥${Utils.moneyShort(total)}</text>`;
  svg.innerHTML = out;

  // 图例表：按金额从大到小
  const sorted = d.items.slice().sort((a, b) => b.value - a.value);
  $('pieLegend').innerHTML = `<tbody>${sorted.map(i => `
    <tr data-id="${i.id}"><td><i style="background:${i.color}"></i>${i.label}</td>
      <td class="r">¥${Utils.money(i.value)}</td>
      <td class="r">${Utils.percent(i.value / total)}</td></tr>`).join('')}
    <tr><td>合计 · ${d.label}</td><td class="r">¥${Utils.money(total)}</td><td class="r">100%</td></tr></tbody>`
    + (d.revenue > 0 ? `<caption style="caption-side:bottom;text-align:left;padding-top:8px" class="src">成本费用占营收 ${Utils.percent(total / d.revenue)}</caption>` : '');

  // 悬停提示
  const tip = $('pieTip');
  svg.querySelectorAll('.slice').forEach(p => {
    p.addEventListener('mouseenter', () => {
      const i = d.items.find(x => x.id === p.dataset.id);
      tip.innerHTML = `<div class="t">${i.label}</div>
        <div class="r"><span><i style="background:${i.color}"></i>${d.label}</span><b>¥${Utils.money(i.value)}</b></div>
        <div class="n">占成本 ${Utils.percent(i.value / total)}${d.revenue > 0 ? '，占营收 ' + Utils.percent(i.value / d.revenue) : ''}</div>`;
      tip.hidden = false;
      tip.style.left = '8px';
      tip.style.top = '8px';
      p.setAttribute('opacity', '0.85');
    });
    p.addEventListener('mouseleave', () => { tip.hidden = true; p.removeAttribute('opacity'); });
  });
}

/* -----------------------------------------------------------------
 * 六、月度利润趋势（柱状图，可点击选月份）
 * ----------------------------------------------------------------- */
function renderBars() {
  const months = plMonths();
  const svg = $('barChart');
  const wrap = $('barWrap');
  const W = Math.max(320, wrap.clientWidth - 28);
  const H = W < 560 ? 260 : 300;
  const pad = { l: 60, r: 12, t: 16, b: 40 };
  const iw = W - pad.l - pad.r, ih = H - pad.t - pad.b;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  if (!months.length) { svg.innerHTML = `<text x="${W / 2}" y="${H / 2}" text-anchor="middle">还没有可计算利润的月份</text>`; return; }

  const pls = months.map(computePL);
  const vals = pls.map(p => p.net);
  const ticks = Utils.niceTicks(Math.min(0, ...vals), Math.max(0, ...vals));
  const lo = ticks[0], hi = ticks[ticks.length - 1];
  const y = v => pad.t + ih - ((v - lo) / (hi - lo)) * ih;
  const band = iw / months.length;
  const bw = Math.max(6, Math.min(36, band * 0.62));
  const x = i => pad.l + band * i + band / 2;

  let out = '';
  ticks.forEach(t => {
    out += `<line x1="${pad.l}" x2="${pad.l + iw}" y1="${y(t)}" y2="${y(t)}" stroke="${t === 0 ? 'var(--muted)' : 'var(--grid)'}" stroke-width="1"/>`;
    out += `<text x="${pad.l - 8}" y="${y(t) + 4}" text-anchor="end">${Utils.moneyShort(t)}</text>`;
  });
  pls.forEach((p, i) => {
    const v = p.net;
    const top = y(Math.max(v, 0)), bottom = y(Math.min(v, 0));
    const h = Math.max(1, bottom - top);
    const rr = Math.min(4, h / 2, bw / 2);
    const x0 = x(i) - bw / 2;
    // 只把远离零线的一端做成圆角
    const path = v >= 0
      ? `M${x0},${bottom}V${top + rr}Q${x0},${top} ${x0 + rr},${top}H${x0 + bw - rr}Q${x0 + bw},${top} ${x0 + bw},${top + rr}V${bottom}Z`
      : `M${x0},${top}V${bottom - rr}Q${x0},${bottom} ${x0 + rr},${bottom}H${x0 + bw - rr}Q${x0 + bw},${bottom} ${x0 + bw},${bottom - rr}V${top}Z`;
    const sel = p.month === state.month;
    out += `<path d="${path}" fill="${v >= 0 ? 'var(--profit-pos)' : 'var(--profit-neg)'}" ${p.partial ? 'fill-opacity="0.45"' : ''}
      ${sel ? 'stroke="var(--fg)" stroke-width="2"' : ''}/>`;
    const mm = Number(p.month.slice(5));
    if (band >= 22 || mm % 3 === 1) out += `<text x="${x(i)}" y="${H - 22}" text-anchor="middle">${mm}</text>`;
    if (mm === 1 || i === 0) out += `<text x="${x(i)}" y="${H - 6}" text-anchor="middle" style="font-weight:600">${p.month.slice(0, 4)}</text>`;
  });
  pls.forEach((p, i) => {
    out += `<rect class="hit" data-i="${i}" x="${pad.l + band * i}" y="${pad.t}" width="${band}" height="${ih}" fill="transparent" style="cursor:pointer"/>`;
  });
  svg.innerHTML = out;

  const tip = $('barTip');
  svg.querySelectorAll('.hit').forEach(rect => {
    const p = pls[Number(rect.dataset.i)];
    rect.addEventListener('mouseenter', () => {
      tip.innerHTML = `<div class="t">${Utils.monthCn(p.month)}${p.partial ? `（营收 ${p.rev.days} 天）` : ''}</div>
        <div class="r"><span>总营收</span><b>¥${Utils.money(p.R)}</b></div>
        <div class="r"><span>总成本费用</span><b>¥${Utils.money(p.totalCost)}</b></div>
        <div class="r"><span><i style="background:${p.net >= 0 ? 'var(--profit-pos)' : 'var(--profit-neg)'}"></i>净利润</span><b>¥${Utils.money(p.net)}</b></div>
        <div class="r"><span>净利率</span><b>${p.netMargin === null ? '—' : Utils.percent(p.netMargin)}</b></div>`;
      tip.hidden = false;
      const scale = svg.getBoundingClientRect().width / W;
      const left = (Number(rect.getAttribute('x')) + band / 2) * scale + 14;
      const tw = tip.offsetWidth;
      tip.style.left = (left + tw + 24 > wrap.clientWidth ? left - tw - 24 : left + 10) + 'px';
      tip.style.top = (pad.t * scale + 6) + 'px';
    });
    rect.addEventListener('mouseleave', () => { tip.hidden = true; });
    rect.addEventListener('click', () => { selectMonth(p.month); });
  });
}

/* -----------------------------------------------------------------
 * 七、月度收支汇总表
 * ----------------------------------------------------------------- */
const SRC_LABEL = { history: '利润表', edited: '利润表（已修改）', manual: '手动录入' };

function summaryRows() {
  return allMonths().map(m => ({ m, pl: computePL(m) })).filter(x => x.pl);
}

/* 汇总表的列：顺序和名称与利润表一致 */
const SUMMARY_COLS = [
  { key: 'R1',            label: '主营业务收入', sub: '房费' },
  { key: 'mainCost',      label: '主营业务成本' },
  { key: 'expense',       label: '费用合计' },
  { key: 'mainProfit',    label: '主营业务营业利润', profit: true },
  { key: 'R2',            label: '非主营业务收入' },
  { key: 'nonMainCost',   label: '非主营业务成本' },
  { key: 'nonMainProfit', label: '非主营业务营业利润', profit: true },
  { key: 'opProfit',      label: '主营业务+非主营业务利润', profit: true },
  { key: 'below',         label: '开办费', sub: '含其他支出' },
  { key: 'net',           label: '净利润', profit: true, strong: true }
];

function renderSummary() {
  const rows = summaryRows();
  const byYear = {};
  rows.forEach(x => { (byYear[x.m.slice(0, 4)] = byYear[x.m.slice(0, 4)] || []).push(x.pl); });
  const colCount = SUMMARY_COLS.length + 3;
  const head = `<thead><tr><th>月份</th>${SUMMARY_COLS.map(c =>
      `<th class="r" style="white-space:normal;min-width:7.5em">${c.label}${c.sub ? `<br><span style="font-weight:400">${c.sub}</span>` : ''}</th>`).join('')}
    <th class="r">净利率</th><th>成本来源</th></tr></thead>`;
  const money = (v, c) => `<td class="r money-cell ${c.strong ? 'total' : ''} ${c.profit && v < 0 ? 'neg' : ''}">${v < 0 ? '−' : ''}¥${Utils.money(Math.abs(v))}</td>`;
  const tr = (label, p, src, cls = '') => `
    <tr class="${cls}" ${cls ? '' : `data-month="${p.month}" style="cursor:pointer"`}>
      <td class="date-cell">${label}</td>
      ${SUMMARY_COLS.map(c => money(p[c.key], c)).join('')}
      <td class="r num">${p.netMargin === null ? '—' : Utils.percent(p.netMargin)}</td>
      <td class="src">${src}</td></tr>`;
  let body = '';
  Object.keys(byYear).sort().reverse().forEach(year => {
    const all = byYear[year];
    const list = all.filter(p => !p.partial);   // 年度合计不含营收未录满的月份
    body += `<tr class="year-head"><td colspan="${colCount}">${year} 年</td></tr>`;
    all.slice().reverse().forEach(p => {
      const extra = p.partial ? `<span class="tag warn">营收 ${p.rev.days} 天</span>` : '';
      body += tr(`${Number(p.month.slice(5))} 月${extra}`, p, (SRC_LABEL[p.costSource] || '') + (p.rev.settled ? ' · 营收用结算数' : ''));
    });
    const t = {};
    SUMMARY_COLS.forEach(c => { t[c.key] = Utils.round2(list.reduce((s, p) => s + p[c.key], 0)); });
    const R = Utils.round2(list.reduce((s, p) => s + p.R, 0));
    t.netMargin = R > 0 ? t.net / R : null;
    body += tr(`${year} 年合计`, t, `${list.length} 个月${all.length > list.length ? '（不含未录满月份）' : ''}`, 'year-row');
  });
  $('summaryTable').innerHTML = head + `<tbody>${body || `<tr><td colspan="${colCount}" class="src">暂无数据</td></tr>`}</tbody>`;
}

/* -----------------------------------------------------------------
 * 八、月度利润表（格式与酒店利润表一致）
 * ----------------------------------------------------------------- */
function statementRows(pl) {
  const rows = [];
  const add = (label, value, lvl, perNight) => rows.push({ label, value, lvl, perNight });
  const group = id => COST_GROUPS.find(g => g.id === id);
  const items = id => group(id).items
    .filter(i => pl.cost[i.key])
    .forEach(i => add(i.label, (i.negative ? -1 : 1) * pl.cost[i.key], 'lvl3', true));
  add('一、主营业务收入（房费）', pl.R1, 'lvl1');
  ROOM_CHANNELS.filter(c => pl.rev[c.key]).forEach(c => add(c.plLabel || c.label, pl.rev[c.key], 'lvl3'));
  add(COST_SECTIONS.mainCost, pl.mainCost, 'lvl1', true);
  add(group('ctrl').label, pl.groups.ctrl, 'lvl2', true); items('ctrl');
  add(group('nonctrl').label, pl.groups.nonctrl, 'lvl2', true); items('nonctrl');
  add(COST_SECTIONS.expense, pl.expense, 'lvl1', true);
  add(group('sales').label, pl.groups.sales, 'lvl2', true); items('sales');
  add(group('admin').label, pl.groups.admin, 'lvl2', true); items('admin');
  add(group('fin').label, pl.groups.fin, 'lvl2', true); items('fin');
  add(group('tax').label, pl.groups.tax, 'lvl2', true);
  add('四、主营业务营业利润（亏损以“－”号填列）', pl.mainProfit, 'result');
  add('五、非主营业务收入', pl.R2, 'lvl1');
  // 小商品POS机 + 发票收入POS机税点 在利润表里合成一行
  NON_ROOM_ITEMS.filter(c => !c.plMerge).forEach(c => {
    const v = Utils.round2(pl.rev[c.key] + NON_ROOM_ITEMS.filter(x => x.plMerge === c.key).reduce((t, x) => t + pl.rev[x.key], 0));
    if (v) add(c.plLabel || c.label, v, 'lvl3');
  });
  add(COST_SECTIONS.nonMainCost, pl.nonMainCost, 'lvl1'); items('nonmain');
  add('七、非主营业务营业利润（亏损以“－”号填列）', pl.nonMainProfit, 'result');
  add('八、主营业务+非主营业务利润', pl.opProfit, 'result');
  add('减：开办费（途泊酒店投建长期费用待摊分摊）', pl.cost.costPreopening, 'lvl2');
  add('减：其他支出', pl.cost.costOtherExp, 'lvl2');
  add('九、净利润', pl.net, 'result');
  return rows;
}

function renderStatement() {
  const pl = computePL(state.month);
  $('stmtTitle').textContent = `月度利润表 · ${Utils.monthCn(state.month)}`;
  if (!pl) {
    $('statement').innerHTML = `<tbody><tr><td class="src" style="padding:24px">${Utils.monthCn(state.month)}缺少营收或成本数据，暂时无法生成利润表。</td></tr></tbody>`;
    return;
  }
  const sold = pl.rev.sold;
  const head = `<thead><tr><td style="font-weight:600">项目</td><td class="r" style="font-weight:600">金额（元）</td>
    <td class="r" style="font-weight:600">间/夜成本</td></tr></thead>`;
  const body = statementRows(pl).map(r => `
    <tr class="${r.lvl}"><td>${r.label}</td>
      <td class="r ${r.lvl === 'result' && r.value < 0 ? 'neg' : ''}">${Utils.money(r.value)}</td>
      <td class="r">${r.perNight && sold > 0 ? Utils.money(r.value / sold) : ''}</td></tr>`).join('');
  const kpi = `<tr class="lvl2"><td>入住间夜 / 入住率 / 平均房价</td><td class="r" colspan="2">${sold.toLocaleString('zh-CN')} 间夜 · ${pl.rev.occ === null ? '—' : Utils.percent(pl.rev.occ)} · ${pl.rev.adr === null ? '—' : '¥' + Utils.money(pl.rev.adr)}</td></tr>
    <tr class="lvl2"><td>毛利 / 毛利率</td><td class="r">${Utils.money(pl.gross)}</td><td class="r">${pl.grossMargin === null ? '' : Utils.percent(pl.grossMargin)}</td></tr>`;
  $('statement').innerHTML = head + `<tbody>${kpi}${body}</tbody>`;
}

/* -----------------------------------------------------------------
 * 九、月度成本录入
 * ----------------------------------------------------------------- */
const costInputs = {};

function buildCostForm() {
  let lastSection = null;
  $('costGroups').innerHTML = COST_GROUPS.map(g => {
    const title = g.section !== lastSection && g.section !== 'nonMainCost' && g.section !== 'below'
      ? `<div class="group-title" style="margin-top:4px">${COST_SECTIONS[g.section]}</div>` : '';
    lastSection = g.section;
    return title + `
    <div class="group" data-group="${g.id}">
      <div class="group-head"><span>${g.label}</span><b id="gsum_${g.id}">¥0.00</b></div>
      <div class="field-row three">
        ${g.items.map(i => `
          <div class="field">
            <label for="c_${i.key}">${i.label}${i.hint ? `<span class="hint">${i.hint}</span>` : ''}</label>
            <div class="money"><input class="input" type="number" id="c_${i.key}" min="0" step="0.01" inputmode="decimal" placeholder="0.00"></div>
            <p class="field-error" data-error-for="c_${i.key}"></p>
          </div>`).join('')}
      </div>
    </div>`;
  }).join('');
  COST_KEYS.forEach(k => { costInputs[k] = $('c_' + k); costInputs[k].addEventListener('input', updateCostTotals); });
}

function loadCostForm() {
  const month = $('cMonth').value;
  const cost = month ? CostStore.get(month) : null;
  COST_KEYS.forEach(k => { costInputs[k].value = cost && cost[k] ? cost[k] : ''; setErr(costInputs[k], ''); });
  const src = month ? CostStore.source(month) : null;
  $('costSource').textContent = !month ? '' : (src === 'history' ? '来自利润表，可修改'
    : src === 'edited' ? '已修改（原数据来自利润表）' : src === 'manual' ? '已录入' : '还没有录入');
  $('btnRestore').hidden = src !== 'edited';
  $('btnCostSave').textContent = src ? '保存修改' : '保存成本';
  updateCostTotals();
}

function readCostForm() {
  const values = {};
  let ok = true, first = null;
  COST_KEYS.forEach(k => {
    const res = Utils.readAmount(costInputs[k]);
    setErr(costInputs[k], res.ok ? '' : res.msg);
    if (!res.ok) { ok = false; if (!first) first = costInputs[k]; } else values[k] = res.value;
  });
  if (!ok) { first.focus(); return null; }
  return values;
}

function updateCostTotals() {
  const v = {};
  COST_KEYS.forEach(k => { const n = Number(costInputs[k].value); v[k] = Number.isFinite(n) && n > 0 ? n : 0; });
  COST_GROUPS.forEach(g => { $('gsum_' + g.id).textContent = '¥' + Utils.money(groupSum(v, g.id)); });
  const total = COST_GROUPS.reduce((s, g) => s + groupSum(v, g.id), 0);
  $('costTotal').textContent = '¥' + Utils.money(total);
}

function setErr(input, msg) {
  const p = document.querySelector(`[data-error-for="${input.id}"]`);
  if (p) p.textContent = msg || '';
  input.classList.toggle('has-error', !!msg);
}

function handleCostSave(e) {
  e.preventDefault();
  const month = $('cMonth').value;
  if (!month) { showToast('请先选择成本月份', true); return; }
  const values = readCostForm();
  if (!values) return;
  const ok = CostStore.save(month, values);
  state.month = month;
  loadCostForm();
  renderAll();
  showToast(ok ? `已保存 ${Utils.monthCn(month)} 成本` : '浏览器无法保存，请检查是否在无痕模式', !ok);
}

function handleCopyPrev() {
  const month = $('cMonth').value;
  if (!month) return;
  const prev = CostStore.get(Utils.prevMonth(month));
  if (!prev) { showToast(`${Utils.monthCn(Utils.prevMonth(month))}没有成本数据可复制`, true); return; }
  COST_KEYS.forEach(k => { costInputs[k].value = prev[k] ? prev[k] : ''; });
  updateCostTotals();
  showToast(`已填入${Utils.monthCn(Utils.prevMonth(month))}的数据，核对修改后请点「保存」`);
}

function handleRestore() {
  const month = $('cMonth').value;
  CostStore.remove(month);
  loadCostForm();
  renderAll();
  showToast(`${Utils.monthCn(month)}已恢复为利润表原数`);
}

/* -----------------------------------------------------------------
 * 十、导出
 * ----------------------------------------------------------------- */
function exportStatement() {
  const pl = computePL(state.month);
  if (!pl) { showToast('这个月缺少营收或成本，无法导出利润表', true); return; }
  if (typeof XLSX === 'undefined') { showToast('Excel 组件没有加载成功，请检查网络后刷新页面再试', true); return; }
  const sold = pl.rev.sold;
  const aoa = [
    [`${Utils.monthCn(state.month)}${CONFIG.HOTEL_NAME}利润表`],
    [`${CONFIG.HOTEL_NAME}（${CONFIG.BRANCH_NAME}）`, '', '单位：元'],
    ['项目名称', '金额', '每间夜'],
    ['可售间夜', pl.rev.avail, ''], ['入住间夜', sold, ''], ['入住率', pl.rev.occ, ''],
    ['平均房价', pl.rev.adr === null ? '' : Utils.round2(pl.rev.adr), ''], ['RevPAR', pl.rev.revpar === null ? '' : Utils.round2(pl.rev.revpar), '']
  ];
  statementRows(pl).forEach(r => aoa.push([(r.lvl === 'lvl3' ? '    ' : '') + r.label, r.value, r.perNight && sold > 0 ? Utils.round2(r.value / sold) : '']));
  aoa.push(['毛利', pl.gross, ''], ['毛利率', pl.grossMargin, ''], ['净利率', pl.netMargin, '']);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  ws['!cols'] = [{ wch: 34 }, { wch: 16 }, { wch: 12 }];
  formatSheetNumbers(ws, 6, aoa.length - 1, 1, 2, []);
  [5, aoa.length - 2, aoa.length - 1].forEach(r => { const c = ws[XLSX.utils.encode_cell({ r, c: 1 })]; if (c && typeof c.v === 'number') c.z = '0.0%'; });
  const c4 = ws[XLSX.utils.encode_cell({ r: 3, c: 1 })]; if (c4) c4.z = '#,##0';
  const c5 = ws[XLSX.utils.encode_cell({ r: 4, c: 1 })]; if (c5) c5.z = '#,##0';
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, Utils.monthCn(state.month));
  const [y, m] = state.month.split('-');
  saveWorkbook(wb, `${CONFIG.HOTEL_NAME}利润表_${y}年${m}月.xlsx`);
}

function exportSummary() {
  if (typeof XLSX === 'undefined') { showToast('Excel 组件没有加载成功，请检查网络后刷新页面再试', true); return; }
  const rows = summaryRows();
  if (!rows.length) { showToast('暂无数据', true); return; }
  // 表一：月度汇总；表二：每月成本明细（也可作为成本数据备份）
  const a = [['月份', ...SUMMARY_COLS.map(c => c.label + (c.sub ? `（${c.sub}）` : '')), '净利率', '毛利率', '营收天数', '成本来源']];
  rows.forEach(({ pl: p }) => a.push([p.month, ...SUMMARY_COLS.map(c => p[c.key]), p.netMargin, p.grossMargin,
    `${p.rev.days}/${p.rev.fullDays}`, SRC_LABEL[p.costSource] || '']));
  const ws = XLSX.utils.aoa_to_sheet(a);
  ws['!cols'] = a[0].map((h, i) => ({ wch: i === 0 ? 10 : 16 }));
  const nCols = SUMMARY_COLS.length;
  formatSheetNumbers(ws, 1, a.length - 1, 1, nCols + 2, [nCols + 1, nCols + 2]);

  const b = [['月份', ...COST_ITEMS.map(i => i.label)]];
  CostStore.months().forEach(m => { const c = CostStore.get(m); b.push([m, ...COST_KEYS.map(k => c[k])]); });
  const ws2 = XLSX.utils.aoa_to_sheet(b);
  ws2['!cols'] = b[0].map(() => ({ wch: 12 }));
  formatSheetNumbers(ws2, 1, b.length - 1, 1, COST_KEYS.length, []);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '月度收支汇总');
  XLSX.utils.book_append_sheet(wb, ws2, '月度成本明细');
  saveWorkbook(wb, `${CONFIG.HOTEL_NAME}月度收支汇总_${Utils.today()}.xlsx`);
}

/* -----------------------------------------------------------------
 * 十一、刷新与初始化
 * ----------------------------------------------------------------- */
function renderImportInfo() {
  const info = ImportStore.info();
  if (!info.records.length) { $('importInfo').textContent = '还没有导入前台 Excel。'; return; }
  const months = {};
  info.records.forEach(r => { months[r.date.slice(0, 7)] = (months[r.date.slice(0, 7)] || 0) + 1; });
  $('importInfo').textContent = '已导入：' + Object.keys(months).sort().map(m => `${Utils.monthCn(m)} ${months[m]} 天`).join('；');
}

function selectMonth(m) {
  state.month = m;
  $('cMonth').value = m;
  loadCostForm();
  renderAll();
}

function renderAll() {
  renderMonthSelect();
  renderKpis();
  renderPie();
  renderBars();
  renderSummary();
  renderStatement();
  renderSettle();
  renderImportInfo();
  $('pieScope').querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', b.dataset.scope === state.pieScope));
}

function init() {
  fillHeader();
  $('storageBanner').hidden = Utils.storageAvailable();
  LedgerStore.load();
  CostStore.load();
  ImportStore.load();
  SettleStore.load();
  buildCostForm();

  renderMonthSelect();
  $('cMonth').value = state.month;
  loadCostForm();

  $('monthSelect').addEventListener('change', e => selectMonth(e.target.value));
  $('pieScope').addEventListener('click', e => {
    const b = e.target.closest('button[data-scope]');
    if (!b) return;
    state.pieScope = b.dataset.scope;
    renderAll();
  });
  $('summaryTable').addEventListener('click', e => {
    const tr = e.target.closest('tr[data-month]');
    if (tr) selectMonth(tr.dataset.month);
  });
  $('cMonth').addEventListener('change', loadCostForm);
  $('costForm').addEventListener('submit', handleCostSave);
  $('btnCopyPrev').addEventListener('click', handleCopyPrev);
  $('btnRestore').addEventListener('click', handleRestore);
  $('btnExportStmt').addEventListener('click', exportStatement);
  $('btnExportSummary').addEventListener('click', exportSummary);
  $('importFile').addEventListener('change', handleImport);
  $('settleForm').addEventListener('submit', handleSettleSave);
  $('btnSettleRestore').addEventListener('click', () => {
    SettleStore.remove(state.month);
    renderAll();
    showToast(`${Utils.monthCn(state.month)}已恢复为营收表原数`);
  });
  $('settleHistory').addEventListener('click', e => {
    const tr = e.target.closest('tr[data-month]');
    if (tr) selectMonth(tr.dataset.month);
  });
  $('btnClearImport').addEventListener('click', () => { ImportStore.clear(); renderAll(); showToast('已清除导入的数据'); });

  let t = null;
  window.addEventListener('resize', () => { clearTimeout(t); t = setTimeout(renderBars, 150); });

  renderAll();
}

init();
