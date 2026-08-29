// trajectory_splitter — frontend logic.
// Talks to Python backend via window.pywebview.api.* (auto-injected by pywebview
// AFTER the pywebviewready event fires — we look it up lazily, not at script load).
// Falls back to a friendly warning if the bridge never arrives
// (e.g. opened directly in a plain browser via file://).

function getApi() {
  return window.pywebview && window.pywebview.api ? window.pywebview.api : null;
}

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const status = $("#status");
let backendReady = false;

function appendLog(line, cls = "") {
  const el = document.createElement("div");
  el.className = "line " + cls;
  el.textContent = line;
  status.appendChild(el);
  status.scrollTop = status.scrollHeight;
}

function nowStamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// 把按钮锁住 -> 调用后端 -> 还原; 出错打 [ERR], 成功打 [OK]
async function runWithButton(btn, busyText, fn) {
  if (!btn) return;
  const original = btn.textContent;
  btn.disabled = true;
  btn.textContent = busyText;
  try {
    return await fn();
  } finally {
    btn.disabled = false;
    btn.textContent = original;
  }
}

// 后端 API 包装 (懒查找, 等 pywebviewready 之后才连得上)
async function withApiCall(label, fn) {
  const a = getApi();
  if (!a) {
    appendLog(`[ERR] ${label} 需要 pywebview 后端`, "err");
    return { ok: false, error: "无后端" };
  }
  return await fn(a);
}

// ── Default POI helpers (split tab) ──────────────────────────────
const DEFAULT_POI = {
  日前: { start: "", end: "【麻涌卸货地址】" },
  后半: { start: "【麻涌卸货地址】", end: "【广州朴朴】" },
};

function refreshPoiDefaults() {
  const dir = document.querySelector('input[name="direction"]:checked').value;
  const d = DEFAULT_POI[dir];
  $("#split-start-poi").value = d.start;
  $("#split-end-poi").value = d.end;
  $("#split-start-poi").placeholder = d.start || "（留空 = 从文件头）";
  $("#split-end-poi").placeholder = d.end;
  $("#start-poi-hint").textContent =
    dir === "日前" ? "日前留空 = 从文件头" : "后半 = 麻涌卸货地址";
  $("#end-poi-hint").textContent =
    dir === "日前" ? "日前 = 麻涌卸货点" : "后半 = 广州朴朴";
}

// ── File/dir pickers ─────────────────────────────────────────────
async function pickFile() {
  return await withApiCall("文件选择", (a) => a.pick_xls_file());
}

async function pickDir() {
  return await withApiCall("目录选择", (a) => a.pick_directory());
}

// ── Split tab ────────────────────────────────────────────────────
function wireSplitTab() {
  $("#pick-split-src").addEventListener("click", async () => {
    const path = await pickFile();
    if (path) {
      $("#split-src").value = path;
      $("#split-summary").classList.remove("visible");
    }
  });

  $("#pick-split-out").addEventListener("click", async () => {
    const path = await pickDir();
    if (path) $("#split-out").value = path;
  });

  $$('input[name="direction"]').forEach((r) =>
    r.addEventListener("change", refreshPoiDefaults)
  );
  refreshPoiDefaults();

  $("#btn-preview").addEventListener("click", async () => {
    const src = $("#split-src").value.trim();
    if (!src) {
      appendLog("[WARN] 先选输入文件", "err");
      return;
    }
    appendLog(`[${nowStamp()}] 预览 ${src}`, "dim");
    const r = await withApiCall("预览主要日期", (a) =>
      a.preview_main_date(src)
    );
    if (!r.ok) {
      appendLog(`[ERR] ${r.error}`, "err");
      $("#split-summary").classList.remove("visible");
      return;
    }
    $("#split-summary-date").textContent = r.main_date;
    $("#split-summary-count").textContent = `(${r.count} 条记录)`;
    $("#split-summary").classList.add("visible");
    appendLog(`主要日期: ${r.main_date} (${r.count} 条)`, "ok");
  });

  $("#btn-run").addEventListener("click", async () => {
    const params = {
      src: $("#split-src").value.trim(),
      direction: document.querySelector('input[name="direction"]:checked').value,
      out_dir: $("#split-out").value.trim(),
      start_poi: $("#split-start-poi").value.trim(),
      end_poi: $("#split-end-poi").value.trim(),
      noon: $("#split-noon").value.trim(),
      clamp: $("#split-clamp").checked,
      round_to_minute: $("#split-round-to-minute").checked,
      tpl: $("#split-tpl").value.trim(),
    };
    if (!params.src || !params.out_dir) {
      appendLog("[WARN] 源文件 / 输出目录都得填", "err");
      return;
    }
    if (!params.end_poi) {
      appendLog("[WARN] end_poi 不能为空", "err");
      return;
    }

    await runWithButton($("#btn-run"), "运行中...", async () => {
      appendLog(
        `[${nowStamp()}] 开始切片 dir=${params.direction} out=${params.out_dir}`,
        "info"
      );
      const r = await withApiCall("切片", (a) => a.run_split(params));
      if (!r.ok) {
        appendLog(`[ERR] ${r.error}`, "err");
        return;
      }
      appendLog(`已输出: ${r.out_path}`, "ok");
    });
  });
}

// ── Merge tab ────────────────────────────────────────────────────
function wireMergeTab() {
  $("#pick-merge-first").addEventListener("click", async () => {
    const path = await pickFile();
    if (path) $("#merge-first").value = path;
  });

  $("#pick-merge-second").addEventListener("click", async () => {
    const path = await pickFile();
    if (path) $("#merge-second").value = path;
  });

  $("#pick-merge-out").addEventListener("click", async () => {
    const path = await pickDir();
    if (path) $("#merge-out").value = path;
  });

  $("#btn-run-merge").addEventListener("click", async () => {
    const params = {
      first: $("#merge-first").value.trim(),
      second: $("#merge-second").value.trim(),
      out_dir: $("#merge-out").value.trim(),
      gap_minutes: $("#merge-gap").value.trim(),
      output_plate: $("#merge-output-plate").value.trim(),
      clamp: $("#merge-clamp").checked,
      round_to_minute: $("#merge-round-to-minute").checked,
      tpl: $("#merge-tpl").value.trim(),
    };
    if (!params.first || !params.second || !params.out_dir) {
      appendLog("[WARN] 前半 / 后半 / 输出目录都得填", "err");
      return;
    }
    await runWithButton($("#btn-run-merge"), "拼接中...", async () => {
      appendLog(
        `[${nowStamp()}] 开始拼接 first=${params.first} second=${params.second}`,
        "info"
      );
      const r = await withApiCall("拼接", (a) => a.run_merge(params));
      if (!r.ok) {
        appendLog(`[ERR] ${r.error}`, "err");
        return;
      }
      appendLog(`已输出: ${r.out_path}`, "ok");
    });
  });
}

// ── Modify tab ───────────────────────────────────────────────────
function wireModifyTab() {
  $("#pick-mod-src").addEventListener("click", async () => {
    const path = await pickFile();
    if (path) $("#mod-src").value = path;
  });

  $("#pick-mod-out").addEventListener("click", async () => {
    const path = await pickDir();
    if (path) $("#mod-out").value = path;
  });

  $("#btn-run-modify").addEventListener("click", async () => {
    const params = {
      src: $("#mod-src").value.trim(),
      out_dir: $("#mod-out").value.trim(),
      output_plate: $("#mod-output-plate").value.trim(),
      time_offset_minutes: $("#mod-offset").value.trim(),
      time_jitter_seconds: $("#mod-time-jitter").value.trim(),
      temp_jitter: $("#mod-temp-jitter").value.trim(),
      clamp: $("#mod-clamp").checked,
      round_to_minute: $("#mod-round-to-minute").checked,
      seed: $("#mod-seed").value.trim(),
      tpl: $("#mod-tpl").value.trim(),
    };
    if (!params.src || !params.out_dir) {
      appendLog("[WARN] 源文件 / 输出目录都得填", "err");
      return;
    }

    await runWithButton($("#btn-run-modify"), "微改中...", async () => {
      appendLog(
        `[${nowStamp()}] 开始微改 src=${params.src}`,
        "info"
      );
      const r = await withApiCall("微改", (a) => a.run_modify(params));
      if (!r.ok) {
        appendLog(`[ERR] ${r.error}`, "err");
        return;
      }
      appendLog(`已输出: ${r.out_path}`, "ok");
    });
  });
}

// ── Tab switching ────────────────────────────────────────────────
function wireTabs() {
  $$("nav.tabs button").forEach((btn) => {
    btn.addEventListener("click", () => {
      const name = btn.dataset.tab;
      if (btn.disabled) return;
      $$("nav.tabs button").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      // 只切换 tab-* 面板, 日志那个 section 不在 tab 切换范围内
      $$("main > section[id^='tab-']").forEach((s) => {
        s.style.display = s.id === `tab-${name}` ? "" : "none";
      });
    });
  });
}

// ── Init ─────────────────────────────────────────────────────────
function init() {
  wireTabs();
  wireSplitTab();
  wireMergeTab();
  wireModifyTab();
  appendLog("trajectory_splitter 已启动", "dim");

  // pywebview 会在页面 load 完成后注入 window.pywebview 并触发 pywebviewready。
  const onReady = () => {
    if (getApi()) {
      backendReady = true;
      appendLog("pywebview 后端已连接 (WebView2)", "ok");
      return;
    }
  };
  window.addEventListener("pywebviewready", onReady);

  // 如果已经被注入但事件已发过, 立即取一次
  if (getApi()) onReady();

  // 1.5 秒还没就绪 -> 大概率是纯浏览器打开, 给提示
  setTimeout(() => {
    if (!backendReady) {
      appendLog(
        "[WARN] 未检测到 pywebview 后端 — 在纯浏览器里只支持日志查看，运行相关按钮需通过 .bat 启动。",
        "err"
      );
    }
  }, 1500);
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}
