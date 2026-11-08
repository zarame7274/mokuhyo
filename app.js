(() => {
  "use strict";

  const API = ((window.APP_CONFIG || {}).API_URL || "").trim();
  const DEMO = !API;
  const WD = ["日", "月", "火", "水", "木", "金", "土"];
  const WD_MON_FIRST = ["月", "火", "水", "木", "金", "土", "日"];
  const POLL_MS = 60_000;

  const $ = (id) => document.getElementById(id);
  const yen = (n) => "¥" + Math.round(n).toLocaleString("ja-JP");
  const qty = (n) => (Math.round(n * 10) / 10).toLocaleString("ja-JP", { maximumFractionDigits: 1 }) + "本";
  const signed = (fmt) => (n) => (n >= 0 ? "+" : "−") + fmt(Math.abs(n));
  const parseNum = (s) => {
    const t = String(s).replace(/[,，¥円本\s]/g, "").replace(/[０-９．]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
    return t === "" ? null : Number(t);
  };
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const todayStr = () => ymd(new Date());
  const parseYmd = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
  const dayLabel = (s) => { const d = parseYmd(s); return `${d.getMonth() + 1}/${d.getDate()}`; };
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  // 売上と本数は同じ仕組みで別々に目標を計算する
  const METRICS = [
    { key: "sales", goalKey: "goalSales", label: "売上", fmt: yen, signed: signed(yen) },
    { key: "qty", goalKey: "goalQty", label: "本数", fmt: qty, signed: signed(qty) },
  ];

  const storage = {
    get(k, fallback) { try { const v = localStorage.getItem(k); return v === null ? fallback : JSON.parse(v); } catch { return fallback; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* 保存できなくても動作は続ける */ } },
  };

  let state = null;
  let selectedDate = null;
  let busy = false;

  // ---------------- 通信 ----------------

  const ERRORS = {
    passcode: "パスコードが違います。",
    no_passcode_configured: "管理者がまだパスコードを設定していません（Apps Script のスクリプト プロパティに PASSCODE を追加してください）。",
    unknown_action: "Apps Script が古いままです。最新の Code.gs を貼り付けて、新しいバージョンでデプロイし直してください。",
  };

  // 読み込み（body なし）も含めてすべて POST。パスコードを URL に載せないため。
  async function api(body) {
    if (DEMO) return demoApi(body);
    const res = await fetch(API, {
      method: "POST",
      // text/plain にすると Apps Script へのクロスオリジン通信で事前確認が不要になる
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ ...(body || { action: "get" }), passcode: storage.get("stt.pass", "") }),
    });
    if (!res.ok) throw new Error(`通信エラー（${res.status}）`);
    const data = await res.json();
    if (data.error) {
      const err = new Error(ERRORS[data.error] || data.error);
      err.code = data.error;
      throw err;
    }
    if (data.goalSales === undefined) {
      throw new Error(ERRORS.unknown_action);
    }
    return data;
  }

  // デモ用：Apps Script と同じ形のデータをこの端末の中だけで作る
  function demoApi(body) {
    const now = new Date();
    let db = storage.get("stt.demo", null);
    if (db && db.unitPrice !== undefined) { // 旧形式（単価×本数）からの引き継ぎ
      db = { ...db, goalSales: db.unitPrice * db.targetQty, goalQty: db.targetQty, qty: {} };
      delete db.unitPrice; delete db.targetQty;
    }
    db = db || {
      goalSales: 3000000, goalQty: 1000,
      month: `${now.getFullYear()}-${pad(now.getMonth() + 1)}`,
      weights: [1, 1, 1, 1, 1, 1, 1], sales: {}, qty: {}, holidays: {},
    };
    if (body) {
      if (body.action === "setEntry") {
        for (const k of ["sales", "qty"]) {
          if (body[k] === undefined) continue;
          if (body[k] === null) delete db[k][body.date]; else db[k][body.date] = body[k];
        }
      } else if (body.action === "setHoliday") {
        if (body.holiday) db.holidays[body.date] = true; else delete db.holidays[body.date];
      } else if (body.action === "setSettings") {
        Object.assign(db, { goalSales: body.goalSales, goalQty: body.goalQty, weights: body.weights });
      } else if (body.action === "startMonth") {
        Object.assign(db, { month: body.month, sales: {}, qty: {}, holidays: {} });
      }
    }
    storage.set("stt.demo", db);
    const [y, m] = db.month.split("-").map(Number);
    const days = [];
    for (let d = new Date(y, m - 1, 1); d.getMonth() === m - 1; d.setDate(d.getDate() + 1)) {
      const key = ymd(d);
      const holiday = !!db.holidays[key];
      days.push({
        date: key, holiday,
        weight: holiday ? 0 : db.weights[(d.getDay() + 6) % 7],
        sales: key in db.sales ? db.sales[key] : null,
        qty: key in db.qty ? db.qty[key] : null,
      });
    }
    return Promise.resolve({ goalSales: db.goalSales, goalQty: db.goalQty, month: db.month, weights: db.weights, days, updatedAt: new Date().toISOString() });
  }

  // ---------------- 計算（スプレッドシートの数式と同じ考え方） ----------------

  // 「これからの日」= 今日以降で、まだ実績を入れていない日。
  // 過去の未入力日は 0 として扱い、残り目標を割り振らない（入れ忘れで目標が甘くならないように）。
  function computeMetric(days, key, goal, today) {
    const total = days.reduce((a, d) => a + (d[key] ?? 0), 0);
    const remaining = Math.max(0, goal - total);
    const isOpen = (d) => d[key] === null && d.date >= today;
    const open = days.filter(isOpen);
    const openWeight = open.reduce((a, d) => a + d.weight, 0);
    const perDay = openWeight ? remaining / openWeight : 0;
    const rows = days.map((d, i) => {
      let target;
      if (!d.weight) target = 0;
      else if (!isOpen(d)) {
        // 実績がある日・過ぎた日：その日の朝の時点で見えていた目標
        const before = days.slice(0, i).reduce((a, x) => a + (x[key] ?? 0), 0);
        const weightFrom = days.slice(i).reduce((a, x) => a + x.weight, 0);
        target = weightFrom ? Math.max(0, goal - before) * d.weight / weightFrom : 0;
      } else {
        // まだの日：今の残り目標を重みで配分
        target = perDay * d.weight;
      }
      return { target, actual: d[key], missed: !isOpen(d) && d[key] === null, diff: d[key] === null ? null : d[key] - target };
    });
    return {
      goal, total, remaining, perDay, rows,
      rate: goal ? total / goal : 0,
      openDays: open.filter((d) => d.weight > 0).length,
    };
  }

  function compute(s) {
    const out = { days: s.days };
    const today = todayStr();
    for (const m of METRICS) out[m.key] = computeMetric(s.days, m.key, s[m.goalKey], today);
    return out;
  }

  // ---------------- 表示 ----------------

  function render() {
    if (!state) return;
    const c = compute(state);
    const today = todayStr();
    const [y, m] = state.month.split("-").map(Number);
    $("monthLabel").textContent = `${y}年${m}月`;

    // 今日の目標（入力済みなら実績）
    const ti = c.days.findIndex((d) => d.date === today);
    $("heroLabel").textContent = ti >= 0
      ? `今日 ${m}月${parseYmd(today).getDate()}日（${WD[parseYmd(today).getDay()]}）`
      : "次の営業日";
    for (const mt of METRICS) {
      const r = c[mt.key];
      const box = $("hero-" + mt.key);
      const todayRow = ti >= 0 ? r.rows[ti] : null;
      const nextIdx = c.days.findIndex((d, i) => d.date >= today && d.weight > 0 && r.rows[i].actual === null);
      let big, sub;
      if (todayRow && todayRow.actual !== null) {
        big = mt.fmt(todayRow.actual);
        sub = `実績 ・ 目標 ${mt.fmt(todayRow.target)}（${mt.signed(todayRow.diff)}）`;
      } else if (nextIdx >= 0) {
        big = mt.fmt(r.rows[nextIdx].target);
        sub = c.days[nextIdx].date === today ? "目標" : `目標（${dayLabel(c.days[nextIdx].date)}）`;
      } else {
        big = mt.fmt(r.total);
        sub = `今月の結果 ・ 達成率 ${(r.rate * 100).toFixed(1)}%`;
      }
      box.querySelector(".big").textContent = big;
      box.querySelector(".sub").textContent = sub;
    }

    // 進み具合
    for (const mt of METRICS) {
      const r = c[mt.key];
      const p = $("prog-" + mt.key);
      p.querySelector(".total").textContent = mt.fmt(r.total);
      p.querySelector(".goal").textContent = mt.fmt(r.goal);
      const pct = Math.min(100, r.rate * 100);
      const bar = p.querySelector(".bar");
      bar.firstElementChild.style.width = pct + "%";
      bar.setAttribute("aria-valuenow", pct.toFixed(0));
      p.querySelector(".rate").textContent = (r.rate * 100).toFixed(1) + "%";
      p.querySelector(".remain").textContent = mt.fmt(r.remaining);
      p.querySelector(".open").textContent = r.openDays + "日";
      p.querySelector(".per").textContent = mt.fmt(r.perDay);
    }

    // 日別テーブル
    $("daysBody").replaceChildren(...c.days.map((d, i) => {
      const tr = document.createElement("tr");
      const wd = parseYmd(d.date).getDay();
      tr.dataset.date = d.date;
      if (d.date === today) tr.classList.add("today");
      if (!d.weight) tr.classList.add("off");
      if (d.date === selectedDate) tr.classList.add("selected");
      const tdDate = document.createElement("td");
      tdDate.innerHTML = `${dayLabel(d.date)} <span class="${wd === 0 ? "sun" : wd === 6 ? "sat" : ""}">${WD[wd]}</span>`;
      tr.append(tdDate);
      for (const mt of METRICS) {
        const r = c[mt.key].rows[i];
        const td = document.createElement("td");
        if (!d.weight && r.actual === null) td.innerHTML = `<span class="muted">休</span>`;
        else {
          const main = r.actual !== null ? esc(mt.fmt(r.actual))
            : r.missed ? `<span class="neg">未入力</span>` : `<span class="muted">—</span>`;
          const diff = r.diff === null ? "" : ` <span class="${r.diff >= 0 ? "pos" : "neg"}">${esc(mt.signed(r.diff))}</span>`;
          td.innerHTML = `<div>${main}</div><div class="small">目標 ${esc(mt.fmt(r.target))}${diff}</div>`;
        }
        tr.append(td);
      }
      const td = document.createElement("td");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = d.holiday;
      cb.disabled = busy || (!d.holiday && d.weight === 0); // 重み0の曜日は設定側で管理
      cb.setAttribute("aria-label", `${dayLabel(d.date)}を休業日にする`);
      cb.addEventListener("click", (e) => e.stopPropagation());
      cb.addEventListener("change", () => send({ action: "setHoliday", date: d.date, holiday: cb.checked }, cb.checked ? "休業日にしました" : "営業日に戻しました"));
      td.append(cb);
      tr.append(td);
      tr.addEventListener("click", () => selectDate(d.date, true));
      return tr;
    }));

    // 入力フォームの日付
    const sel = $("entryDate");
    const keep = selectedDate && c.days.some((d) => d.date === selectedDate) ? selectedDate
      : (c.days.some((d) => d.date === today) ? today : c.days[0]?.date);
    sel.replaceChildren(...c.days.map((d) => {
      const o = document.createElement("option");
      o.value = d.date;
      const done = d.sales !== null || d.qty !== null;
      o.textContent = `${dayLabel(d.date)}（${WD[parseYmd(d.date).getDay()]}）` + (done ? " ✓" : "");
      return o;
    }));
    if (keep) { sel.value = keep; if (selectedDate !== keep) selectDate(keep, false); }
    const selRow = c.days.find((d) => d.date === selectedDate);
    $("clearBtn").hidden = !(selRow && (selRow.sales !== null || selRow.qty !== null));
    updateEntryHint();

    // 設定（入力中は上書きしない）
    if (!$("settings").contains(document.activeElement)) fillSettings();

    const t = new Date(state.updatedAt);
    $("syncLabel").textContent = `${pad(t.getHours())}:${pad(t.getMinutes())} 更新`;
  }

  function selectDate(date, scroll) {
    selectedDate = date;
    $("entryDate").value = date;
    const row = state.days.find((d) => d.date === date);
    $("entrySales").value = row && row.sales !== null ? row.sales : "";
    $("entryQty").value = row && row.qty !== null ? row.qty : "";
    $("clearBtn").hidden = !(row && (row.sales !== null || row.qty !== null));
    document.querySelectorAll("#daysBody tr").forEach((tr) => tr.classList.toggle("selected", tr.dataset.date === date));
    updateEntryHint();
    if (scroll) { $("entryForm").scrollIntoView({ behavior: "smooth", block: "center" }); $("entrySales").focus({ preventScroll: true }); }
  }

  function updateEntryHint() {
    if (!state) return;
    const c = compute(state);
    const i = c.days.findIndex((x) => x.date === $("entryDate").value);
    const parts = [];
    if (i >= 0 && c.days[i].weight) parts.push(`この日の目標 ${yen(c.sales.rows[i].target)} ・ ${qty(c.qty.rows[i].target)}`);
    const s = parseNum($("entrySales").value), q = parseNum($("entryQty").value);
    if (s > 0 && q > 0) parts.push(`平均単価 ${yen(s / q)}`);
    $("entryHint").textContent = parts.join(" ／ ");
  }

  function fillSettings() {
    $("setGoalSales").value = state.goalSales;
    $("setGoalQty").value = state.goalQty;
    const box = $("weightInputs");
    if (!box.children.length) {
      WD_MON_FIRST.forEach((w, i) => {
        const l = document.createElement("label");
        l.textContent = w;
        const inp = document.createElement("input");
        inp.inputMode = "decimal";
        inp.id = "w" + i;
        l.append(inp);
        box.append(l);
      });
    }
    state.weights.forEach((w, i) => { $("w" + i).value = w; });
    const [y, m] = state.month.split("-").map(Number);
    const nm = new Date(y, m, 1);
    if (!$("newMonth").value) $("newMonth").value = `${nm.getFullYear()}-${pad(nm.getMonth() + 1)}`;
    updateGoalPreview();
  }

  function updateGoalPreview() {
    const s = parseNum($("setGoalSales").value), q = parseNum($("setGoalQty").value);
    $("setUnitPreview").textContent = s > 0 && q > 0 ? yen(s / q) : "—";
  }

  // ---------------- 操作 ----------------

  function showError(msg) {
    const box = $("errorBox");
    box.textContent = msg || "";
    box.hidden = !msg;
  }

  function showGate(on) {
    const wasShown = !$("gate").hidden;
    $("gate").hidden = !on;
    $("main").hidden = on;
    if (on) {
      state = null;
      $("monthLabel").textContent = "パスコードが必要です";
      if (!wasShown) $("gateInput").focus();
    }
  }

  let toastTimer;
  function toast(msg) {
    const t = $("toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 2200);
  }

  function setBusy(b) {
    busy = b;
    document.querySelectorAll("#main button, #main select, #main input").forEach((el) => { el.disabled = b; });
    if (!b) render(); // 休業日チェックの有効/無効は render で決め直す
  }

  async function send(body, okMsg) {
    if (busy) return false;
    setBusy(true);
    let ok = false;
    try {
      state = await api(body);
      showError("");
      if (okMsg) toast(okMsg);
      ok = true;
    } catch (e) {
      showError(e.message || "保存できませんでした");
      if (e.code === "passcode") showGate(true);
    } finally {
      setBusy(false);
    }
    return ok;
  }

  async function load() {
    if (busy) return;
    if (!DEMO && !storage.get("stt.pass", "")) { showGate(true); return; }
    try {
      state = await api(null);
      showError("");
      showGate(false);
      render();
    } catch (e) {
      showError("読み込めませんでした：" + (e.message || e));
      if (e.code === "passcode") { storage.set("stt.pass", ""); showGate(true); } // 間違ったパスコードで再送し続けない
      else if (!state) $("monthLabel").textContent = "読み込みに失敗しました";
    }
  }

  function readEntry(id, label) {
    const v = parseNum($(id).value);
    if (v !== null && (Number.isNaN(v) || v < 0)) throw new Error(`${label}は0以上の数字で入力してください`);
    return v;
  }

  $("entryDate").addEventListener("change", (e) => selectDate(e.target.value, false));
  $("entrySales").addEventListener("input", updateEntryHint);
  $("entryQty").addEventListener("input", updateEntryHint);
  $("entryForm").addEventListener("submit", (e) => {
    e.preventDefault();
    let sales, q;
    try { sales = readEntry("entrySales", "売上金額"); q = readEntry("entryQty", "本数"); }
    catch (err) { showError(err.message); return; }
    if (sales === null && q === null) { showError("売上金額か本数を入力してください（消すときは「この日の入力を消す」）"); return; }
    const date = $("entryDate").value;
    send({ action: "setEntry", date, sales, qty: q }, `${dayLabel(date)} の実績を保存しました`);
  });
  $("clearBtn").addEventListener("click", () =>
    send({ action: "setEntry", date: $("entryDate").value, sales: null, qty: null }, "入力を消しました"));

  $("setGoalSales").addEventListener("input", updateGoalPreview);
  $("setGoalQty").addEventListener("input", updateGoalPreview);
  $("settingsForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const goalSales = parseNum($("setGoalSales").value), goalQty = parseNum($("setGoalQty").value);
    const weights = WD_MON_FIRST.map((_, i) => parseNum($("w" + i).value));
    if (!(goalSales > 0) || !(goalQty > 0)) { showError("月間目標売上と月間目標本数は1以上の数字にしてください"); return; }
    if (weights.some((w) => w === null || Number.isNaN(w) || w < 0)) { showError("曜日の重みは0以上の数字にしてください"); return; }
    if (weights.every((w) => w === 0)) { showError("すべての曜日を0にはできません"); return; }
    document.activeElement?.blur();
    send({ action: "setSettings", goalSales, goalQty, weights }, "設定を保存しました");
  });

  $("startMonthBtn").addEventListener("click", () => {
    if (!$("newMonth").value) { showError("始める月を選んでください"); return; }
    $("startMonthConfirm").hidden = false;
  });
  $("startMonthNo").addEventListener("click", () => { $("startMonthConfirm").hidden = true; });
  $("startMonthYes").addEventListener("click", async () => {
    $("startMonthConfirm").hidden = true;
    selectedDate = null;
    if (await send({ action: "startMonth", month: $("newMonth").value }, "新しい月を始めました")) {
      $("newMonth").value = "";
      fillSettings();
    }
  });

  $("gate").addEventListener("submit", (e) => {
    e.preventDefault();
    const v = $("gateInput").value.trim();
    if (!v) return;
    storage.set("stt.pass", v);
    $("gateInput").value = "";
    showError("");
    load();
  });
  $("logoutBtn").addEventListener("click", () => {
    storage.set("stt.pass", "");
    $("settings").open = false;
    showGate(true);
    toast("この端末からパスコードを消しました");
  });

  $("refreshBtn").addEventListener("click", load);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) load(); });
  setInterval(() => { if (!document.hidden) load(); }, POLL_MS);

  $("demoBanner").hidden = !DEMO;
  load();
})();
