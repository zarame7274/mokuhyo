(() => {
  "use strict";

  const API = ((window.APP_CONFIG || {}).API_URL || "").trim();
  const DEMO = !API;
  const WD = ["日", "月", "火", "水", "木", "金", "土"];
  const WD_MON_FIRST = ["月", "火", "水", "木", "金", "土", "日"];
  const POLL_MS = 60_000;

  const $ = (id) => document.getElementById(id);
  const yen = (n) => "¥" + Math.round(n).toLocaleString("ja-JP");
  const signedYen = (n) => (n >= 0 ? "+" : "−") + yen(Math.abs(n));
  const qty = (n) => n.toLocaleString("ja-JP", { maximumFractionDigits: 1 }) + "本";
  const parseNum = (s) => {
    const t = String(s).replace(/[,，¥円\s]/g, "").replace(/[０-９．]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
    return t === "" ? null : Number(t);
  };
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const todayStr = () => ymd(new Date());
  const parseYmd = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
  const dayLabel = (s) => { const d = parseYmd(s); return `${d.getMonth() + 1}/${d.getDate()}`; };

  const storage = {
    get(k, fallback) { try { const v = localStorage.getItem(k); return v === null ? fallback : JSON.parse(v); } catch { return fallback; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* 保存できなくても動作は続ける */ } },
  };

  let state = null;
  let selectedDate = null;
  let busy = false;

  // ---------------- 通信 ----------------

  async function api(body) {
    if (DEMO) return demoApi(body);
    let res;
    if (body) {
      res = await fetch(API, {
        method: "POST",
        // text/plain にすると Apps Script へのクロスオリジン通信で事前確認が不要になる
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ ...body, passcode: storage.get("stt.pass", "") }),
      });
    } else {
      res = await fetch(API + (API.includes("?") ? "&" : "?") + "t=" + Date.now());
    }
    if (!res.ok) throw new Error(`通信エラー（${res.status}）`);
    const data = await res.json();
    if (data.error) {
      const err = new Error(data.error === "passcode" ? "パスコードが違います。設定の一番下で入力してください。" : data.error);
      err.code = data.error;
      throw err;
    }
    return data;
  }

  // デモ用：Apps Script と同じ形のデータをこの端末の中だけで作る
  function demoApi(body) {
    const now = new Date();
    const db = storage.get("stt.demo", null) || {
      unitPrice: 3000, targetQty: 1000,
      month: `${now.getFullYear()}-${pad(now.getMonth() + 1)}`,
      weights: [1, 1, 1, 1, 1, 1, 1], sales: {}, holidays: {},
    };
    if (body) {
      if (body.action === "setSales") {
        if (body.sales === null) delete db.sales[body.date]; else db.sales[body.date] = body.sales;
      } else if (body.action === "setHoliday") {
        if (body.holiday) db.holidays[body.date] = true; else delete db.holidays[body.date];
      } else if (body.action === "setSettings") {
        Object.assign(db, { unitPrice: body.unitPrice, targetQty: body.targetQty, weights: body.weights });
      } else if (body.action === "startMonth") {
        Object.assign(db, { month: body.month, sales: {}, holidays: {} });
      }
      storage.set("stt.demo", db);
    }
    const [y, m] = db.month.split("-").map(Number);
    const days = [];
    for (let d = new Date(y, m - 1, 1); d.getMonth() === m - 1; d.setDate(d.getDate() + 1)) {
      const key = ymd(d);
      const holiday = !!db.holidays[key];
      days.push({
        date: key, holiday,
        weight: holiday ? 0 : db.weights[(d.getDay() + 6) % 7],
        sales: key in db.sales ? db.sales[key] : null,
      });
    }
    return Promise.resolve({ ...db, days, updatedAt: new Date().toISOString() });
  }

  // ---------------- 計算（スプレッドシートの数式と同じ考え方） ----------------

  function compute(s) {
    const goal = s.unitPrice * s.targetQty;
    const days = s.days.map((d) => ({ ...d }));
    const total = days.reduce((a, d) => a + (d.sales ?? 0), 0);
    const remaining = Math.max(0, goal - total);
    const open = days.filter((d) => d.sales === null);
    const openWeight = open.reduce((a, d) => a + d.weight, 0);
    const perDay = openWeight ? remaining / openWeight : 0;

    days.forEach((d, i) => {
      if (!d.weight) d.target = 0;
      else if (d.sales !== null) {
        // 実績がある日：その日の朝の時点で見えていた目標
        const before = days.slice(0, i).reduce((a, x) => a + (x.sales ?? 0), 0);
        const weightFrom = days.slice(i).reduce((a, x) => a + x.weight, 0);
        d.target = weightFrom ? Math.max(0, goal - before) * d.weight / weightFrom : 0;
      } else {
        // まだの日：今の残り目標を重みで配分
        d.target = perDay * d.weight;
      }
      d.diff = d.sales === null ? null : d.sales - d.target;
    });

    return {
      goal, total, remaining, perDay, days,
      rate: goal ? total / goal : 0,
      openDays: open.filter((d) => d.weight > 0).length,
    };
  }

  // ---------------- 表示 ----------------

  function render() {
    if (!state) return;
    const c = compute(state);
    const unit = state.unitPrice || 1;
    const today = todayStr();
    const [y, m] = state.month.split("-").map(Number);
    $("monthLabel").textContent = `${y}年${m}月`;

    // 今日（または次の営業日）の目標
    const todayRow = c.days.find((d) => d.date === today);
    const next = c.days.find((d) => d.date >= today && d.sales === null && d.weight > 0);
    if (todayRow && todayRow.sales !== null) {
      $("heroLabel").textContent = "今日の実績";
      $("heroValue").textContent = yen(todayRow.sales);
      $("heroSub").textContent = `目標 ${yen(todayRow.target)}（${signedYen(todayRow.diff)}）` +
        (next ? ` ・ 次の目標 ${dayLabel(next.date)} ${yen(next.target)}` : "");
    } else if (next) {
      $("heroLabel").textContent = next.date === today ? "今日の目標" : `次の目標（${dayLabel(next.date)}）`;
      $("heroValue").textContent = yen(next.target);
      $("heroSub").textContent = `${qty(next.target / unit)} ・ 単価 ${yen(state.unitPrice)}`;
    } else {
      $("heroLabel").textContent = "今月の結果";
      $("heroValue").textContent = yen(c.total);
      $("heroSub").textContent = `達成率 ${(c.rate * 100).toFixed(1)}%`;
    }

    $("totalValue").textContent = `${yen(c.total)}（${qty(c.total / unit)}）`;
    $("goalValue").textContent = `${yen(c.goal)}（${qty(state.targetQty)}）`;
    const pct = Math.min(100, c.rate * 100);
    $("barFill").style.width = pct + "%";
    $("bar").setAttribute("aria-valuenow", pct.toFixed(0));
    $("rateValue").textContent = (c.rate * 100).toFixed(1) + "%";
    $("remainValue").textContent = yen(c.remaining);
    $("openDaysValue").textContent = c.openDays + "日";
    $("perDayValue").textContent = `${yen(c.perDay)} / ${qty(c.perDay / unit)}`;

    // 日別テーブル
    const body = $("daysBody");
    body.replaceChildren(...c.days.map((d) => {
      const tr = document.createElement("tr");
      const wd = parseYmd(d.date).getDay();
      tr.dataset.date = d.date;
      if (d.date === today) tr.classList.add("today");
      if (!d.weight) tr.classList.add("off");
      if (d.date === selectedDate) tr.classList.add("selected");
      const cells = [
        `${dayLabel(d.date)} <span class="${wd === 0 ? "sun" : wd === 6 ? "sat" : ""}">${WD[wd]}</span>`,
        d.weight ? yen(d.target) : "休",
        d.sales === null ? "" : yen(d.sales),
        d.diff === null ? "" : `<span class="${d.diff >= 0 ? "pos" : "neg"}">${signedYen(d.diff)}</span>`,
      ];
      cells.forEach((html) => { const td = document.createElement("td"); td.innerHTML = html; tr.append(td); });
      const td = document.createElement("td");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = d.holiday;
      cb.disabled = !d.holiday && d.weight === 0; // 重み0の曜日は設定側で管理
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
      o.textContent = `${dayLabel(d.date)}（${WD[parseYmd(d.date).getDay()]}）` + (d.sales !== null ? " ✓" : "");
      return o;
    }));
    if (keep) { sel.value = keep; if (selectedDate !== keep) selectDate(keep, false); }
    const selRow = c.days.find((d) => d.date === selectedDate);
    $("clearBtn").hidden = !(selRow && selRow.sales !== null);
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
    $("clearBtn").hidden = !(row && row.sales !== null);
    document.querySelectorAll("#daysBody tr").forEach((tr) => tr.classList.toggle("selected", tr.dataset.date === date));
    updateEntryHint();
    if (scroll) { $("entryForm").scrollIntoView({ behavior: "smooth", block: "center" }); $("entrySales").focus({ preventScroll: true }); }
  }

  function updateEntryHint() {
    if (!state) return;
    const c = compute(state);
    const d = c.days.find((x) => x.date === $("entryDate").value);
    const v = parseNum($("entrySales").value);
    const unit = state.unitPrice || 1;
    const parts = [];
    if (d && d.weight) parts.push(`この日の目標 ${yen(d.target)}`);
    if (v !== null && !Number.isNaN(v)) parts.push(`入力額 ≒ ${qty(v / unit)}`);
    $("entryHint").textContent = parts.join(" ・ ");
  }

  function fillSettings() {
    $("setUnit").value = state.unitPrice;
    $("setQty").value = state.targetQty;
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
    const u = parseNum($("setUnit").value), q = parseNum($("setQty").value);
    $("setGoalPreview").textContent = u && q ? yen(u * q) : "—";
  }

  // ---------------- 操作 ----------------

  function showError(msg) {
    const box = $("errorBox");
    box.textContent = msg || "";
    box.hidden = !msg;
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
    document.querySelectorAll("button, select, input").forEach((el) => { el.disabled = b; });
    if (!b) render(); // 休業日チェックの有効/無効は render で決め直す
  }

  async function send(body, okMsg) {
    if (busy) return;
    setBusy(true);
    try {
      state = await api(body);
      showError("");
      if (okMsg) toast(okMsg);
    } catch (e) {
      showError(e.message || "保存できませんでした");
      if (e.code === "passcode") $("settings").open = true;
    } finally {
      setBusy(false);
    }
  }

  async function load() {
    if (busy) return;
    try {
      state = await api(null);
      showError("");
      render();
    } catch (e) {
      showError("読み込めませんでした：" + (e.message || e));
      if (!state) $("monthLabel").textContent = "読み込みに失敗しました";
    }
  }

  $("entryDate").addEventListener("change", (e) => selectDate(e.target.value, false));
  $("entrySales").addEventListener("input", updateEntryHint);
  $("entryForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const v = parseNum($("entrySales").value);
    if (v === null || Number.isNaN(v) || v < 0) { showError("売上金額を数字で入力してください"); return; }
    send({ action: "setSales", date: $("entryDate").value, sales: v }, `${dayLabel($("entryDate").value)} の売上を保存しました`);
  });
  $("clearBtn").addEventListener("click", () =>
    send({ action: "setSales", date: $("entryDate").value, sales: null }, "入力を消しました"));

  $("setUnit").addEventListener("input", updateGoalPreview);
  $("setQty").addEventListener("input", updateGoalPreview);
  $("settingsForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const unitPrice = parseNum($("setUnit").value), targetQty = parseNum($("setQty").value);
    const weights = WD_MON_FIRST.map((_, i) => parseNum($("w" + i).value));
    if (!(unitPrice > 0) || !(targetQty > 0)) { showError("単価と目標本数は1以上の数字にしてください"); return; }
    if (weights.some((w) => w === null || Number.isNaN(w) || w < 0)) { showError("曜日の重みは0以上の数字にしてください"); return; }
    if (weights.every((w) => w === 0)) { showError("すべての曜日を0にはできません"); return; }
    document.activeElement?.blur();
    send({ action: "setSettings", unitPrice, targetQty, weights }, "設定を保存しました");
  });

  $("startMonthBtn").addEventListener("click", () => {
    if (!$("newMonth").value) { showError("始める月を選んでください"); return; }
    $("startMonthConfirm").hidden = false;
  });
  $("startMonthNo").addEventListener("click", () => { $("startMonthConfirm").hidden = true; });
  $("startMonthYes").addEventListener("click", async () => {
    $("startMonthConfirm").hidden = true;
    selectedDate = null;
    await send({ action: "startMonth", month: $("newMonth").value }, "新しい月を始めました");
    $("newMonth").value = "";
    if (state) fillSettings();
  });

  $("passInput").value = storage.get("stt.pass", "");
  $("passSave").addEventListener("click", () => { storage.set("stt.pass", $("passInput").value); toast("パスコードを記憶しました"); });

  $("refreshBtn").addEventListener("click", load);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) load(); });
  setInterval(() => { if (!document.hidden) load(); }, POLL_MS);

  $("demoBanner").hidden = !DEMO;
  load();
})();
