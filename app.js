const STORAGE_KEY = "work-tracker-data-v1";
const STORAGE_RECOVERY_KEY = "work-tracker-data-v1-recovery";
const CLOUD_ENDPOINT_KEY = "work-tracker-cloud-endpoint";
const CLOUD_ENDPOINT_DEFAULT =
  "https://script.google.com/macros/s/AKfycbxJZtFGoBFuRPcW-EdD9WMaOE716tFNJFHWYQQY0-hHPzOJOUPZ5pM0SezJye9xf6GmJg/exec";
const AUTO_SAVE_INTERVAL_MS = 15000;
const AUTO_CLOUD_SAVE_INTERVAL_MS = 300000;
const AUTO_CLOUD_LOAD_INTERVAL_MS = 600000;
const AUTO_CLOUD_LATEST_CHECK_INTERVAL_MS = 60000;
const AUTO_RECURRENCE_CHECK_MS = 60000;
const TIMELINE_MIN_EVENT_HEIGHT = 18;
const SPREADSHEET_ID = "1ggWSLbaj5vFMmkcJP4EWAUxQusQ12m8jpWmta0-lmDg";
const APP_VERSION_FALLBACK = "1970-01-01 00:00";
const STATUS_OPTIONS = ["未着手", "着手", "チェック中", "完了", "取り下げ"];
const REPEAT_OPTIONS = ["none", "daily", "weekly", "monthly"];

const state = loadState();
let isCloudSyncBusy = false;
let dragDropdownListenerInitialized = false;
let draggingTaskId = "";
let cloudSaveDebounceTimer = null;
let cloudSavePending = false;
let lastSeenCloudSavedAt = 0;

const el = {
  openTimelineModal: document.getElementById("open-timeline-modal"),
  timelineModal: document.getElementById("timeline-modal"),
  closeTimelineModal: document.getElementById("close-timeline-modal"),
  timelineFilter: document.getElementById("timeline-filter"),
  timelinePrevDay: document.getElementById("timeline-prev-day"),
  timelineToday: document.getElementById("timeline-today"),
  timelineNextDay: document.getElementById("timeline-next-day"),
  timelineViewToggle: document.getElementById("timeline-view-toggle"),
  timelineDayDate: document.getElementById("timeline-day-date"),
  timelineKeyword: document.getElementById("timeline-keyword"),
  timelineMeta: document.getElementById("timeline-meta"),
  timelineHoverDetail: document.getElementById("timeline-hover-detail"),
  timelineList: document.getElementById("timeline-list"),
  addSectionsToggle: document.getElementById("add-sections-toggle"),
  addSectionsPanel: document.getElementById("add-sections-panel"),
  topForm: document.getElementById("top-form"),
  topName: document.getElementById("top-name"),
  largeForm: document.getElementById("large-form"),
  largeTopSelect: document.getElementById("large-top-select"),
  largeName: document.getElementById("large-name"),
  midForm: document.getElementById("mid-form"),
  midLargeSelect: document.getElementById("mid-large-select"),
  midName: document.getElementById("mid-name"),
  taskForm: document.getElementById("task-form"),
  taskMidSelect: document.getElementById("task-mid-select"),
  taskName: document.getElementById("task-name"),
  taskStatus: document.getElementById("task-status"),
  taskTags: document.getElementById("task-tags"),
  manualForm: document.getElementById("manual-form"),
  manualToggle: document.getElementById("manual-toggle"),
  manualPanel: document.getElementById("manual-panel"),
  manualTaskSelect: document.getElementById("manual-task-select"),
  manualDate: document.getElementById("manual-date"),
  manualHours: document.getElementById("manual-hours"),
  manualMinutes: document.getElementById("manual-minutes"),
  backupExport: document.getElementById("backup-export"),
  backupImport: document.getElementById("backup-import"),
  cloudEndpoint: document.getElementById("cloud-endpoint"),
  cloudSave: document.getElementById("cloud-save"),
  cloudLoad: document.getElementById("cloud-load"),
  backupFile: document.getElementById("backup-file"),
  taskLargeFilter: document.getElementById("task-large-filter"),
  taskMidFilter: document.getElementById("task-mid-filter"),
  statusFilter: document.getElementById("status-filter"),
  taskList: document.getElementById("task-list"),
  taskTpl: document.getElementById("task-row-template"),
  activeStatus: document.getElementById("active-status"),
  rangeSelect: document.getElementById("range-select"),
  summaryPrev: document.getElementById("summary-prev"),
  summaryNext: document.getElementById("summary-next"),
  summaryReset: document.getElementById("summary-reset"),
  summaryTabs: document.getElementById("summary-tabs"),
  summaryPeriod: document.getElementById("summary-period"),
  summaryTotal: document.getElementById("summary-total"),
  summaryTask: document.getElementById("summary-task"),
  summaryMid: document.getElementById("summary-mid"),
  summaryLarge: document.getElementById("summary-large"),
  summaryTop: document.getElementById("summary-top"),
  summaryTags: document.getElementById("summary-tags"),
  appVersion: document.getElementById("app-version"),
};

el.cloudEndpoint.value = localStorage.getItem(CLOUD_ENDPOINT_KEY) || CLOUD_ENDPOINT_DEFAULT;
if (!el.manualDate.value) {
  el.manualDate.value = formatDateInput(new Date());
}

bindEvents();
initDragDropdownGlobalClose();
if (applyRecurringResets()) {
  persistState();
}
initializeVersionInfo();
renderAll();
setInterval(() => {
  renderActiveStatus();
  renderSummary();
  if (!el.timelineModal.classList.contains("hidden")) {
    renderTimelineModal();
  }
}, 1000);
setInterval(() => {
  const recurringChanged = applyRecurringResets();
  if (recurringChanged) {
    persistAndRender();
  }
}, AUTO_RECURRENCE_CHECK_MS);
startAutoSave();
startCloudAutoSync();
registerServiceWorker();

function bindEvents() {
  el.openTimelineModal.addEventListener("click", openTimelineModal);
  el.closeTimelineModal.addEventListener("click", closeTimelineModal);
  el.timelineModal.addEventListener("click", (e) => {
    if (e.target.dataset.close === "1") closeTimelineModal();
  });
  el.timelineFilter.addEventListener("submit", (e) => {
    e.preventDefault();
    renderTimelineModal();
  });
  el.timelinePrevDay.addEventListener("click", () => {
    shiftTimelineDay(-1);
  });
  el.timelineToday.addEventListener("click", () => {
    el.timelineDayDate.value = formatDateInput(new Date());
    renderTimelineModal();
  });
  el.timelineNextDay.addEventListener("click", () => {
    shiftTimelineDay(1);
  });
  el.timelineViewToggle.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-view]");
    if (!btn) return;
    state.timelineView = normalizeTimelineView(btn.dataset.view);
    persistUiAndRender();
  });
  el.addSectionsToggle.addEventListener("click", () => {
    state.addSectionsCollapsed = !state.addSectionsCollapsed;
    persistUiAndRender();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !el.timelineModal.classList.contains("hidden")) {
      closeTimelineModal();
    }
  });

  el.topForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const name = el.topName.value.trim();
    if (!name) return;
    state.topGroups.push({ id: uid(), name, updatedAt: Date.now() });
    el.topName.value = "";
    persistAndRender();
  });

  el.largeForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const topGroupId = state.uiSelections.largeTopId;
    const name = el.largeName.value.trim();
    if (!topGroupId || !name) return;
    state.largeGroups.push({ id: uid(), name, topGroupId, updatedAt: Date.now() });
    el.largeName.value = "";
    persistAndRender();
  });

  el.midForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const largeGroupId = state.uiSelections.midLargeId;
    const name = el.midName.value.trim();
    if (!largeGroupId || !name) return;
    state.midGroups.push({ id: uid(), name, largeGroupId, updatedAt: Date.now() });
    el.midName.value = "";
    persistAndRender();
  });

  el.taskForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const parent = parseTaskParentValue(state.uiSelections.taskParentValue);
    const name = el.taskName.value.trim();
    const status = normalizeStatus(el.taskStatus.value);
    const tags = parseTags(el.taskTags.value);
    if (!parent || !name) return;
    state.tasks.push({
      id: uid(),
      name,
      parentType: parent.parentType,
      midGroupId: parent.midGroupId || null,
      largeGroupId: parent.largeGroupId || null,
      status,
      isTodayTask: false,
      recurrence: "none",
      recurrenceResetKey: "",
      tags,
      updatedAt: Date.now(),
    });
    el.taskName.value = "";
    el.taskTags.value = "";
    el.taskStatus.value = "未着手";
    persistAndRender();
  });

  el.manualForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const taskId = el.manualTaskSelect.value;
    const dateText = el.manualDate.value;
    const hours = parseInt(el.manualHours.value, 10);
    const minutes = parseInt(el.manualMinutes.value, 10);
    if (!taskId || !dateText || !Number.isFinite(hours) || !Number.isFinite(minutes)) {
      alert("タスク・日付・時間・分を正しく入力してください。");
      return;
    }
    if (hours < 0 || minutes < 0 || minutes > 59) {
      alert("時間は0以上、分は0〜59で入力してください。");
      return;
    }
    const totalMinutes = hours * 60 + minutes;
    if (totalMinutes <= 0) {
      alert("工数は1分以上入力してください。");
      return;
    }
    const date = new Date(`${dateText}T00:00:00`);
    const startAt = date.getTime();
    const endAt = startAt + totalMinutes * 60000;
    if (!Number.isFinite(startAt) || !Number.isFinite(endAt)) {
      alert("日付を正しく入力してください。");
      return;
    }
    addManualSession(taskId, startAt, endAt);
    el.manualHours.value = "0";
    el.manualMinutes.value = "0";
    persistAndRender();
  });
  el.manualToggle.addEventListener("click", () => {
    state.manualCollapsed = !state.manualCollapsed;
    persistUiAndRender();
  });

  el.rangeSelect.addEventListener("change", () => {
    renderSummary();
  });

  el.summaryPrev.addEventListener("click", () => {
    const range = el.rangeSelect.value;
    const current = state.summaryOffsets[range] || 0;
    state.summaryOffsets[range] = current - 1;
    persistUiAndRender();
  });

  el.summaryNext.addEventListener("click", () => {
    const range = el.rangeSelect.value;
    const current = state.summaryOffsets[range] || 0;
    state.summaryOffsets[range] = Math.min(0, current + 1);
    persistUiAndRender();
  });

  el.summaryReset.addEventListener("click", () => {
    const range = el.rangeSelect.value;
    state.summaryOffsets[range] = 0;
    persistUiAndRender();
  });

  el.summaryTabs.addEventListener("click", (e) => {
    const target = e.target.closest(".summary-tab-btn");
    if (!target) return;
    const tab = target.dataset.tab;
    state.summaryTab = normalizeSummaryTab(tab);
    persistUiAndRender();
  });

  el.backupExport.addEventListener("click", exportBackup);
  el.backupImport.addEventListener("click", () => {
    el.backupFile.click();
  });
  el.backupFile.addEventListener("change", importBackup);
  el.cloudEndpoint.addEventListener("change", () => {
    localStorage.setItem(CLOUD_ENDPOINT_KEY, el.cloudEndpoint.value.trim());
  });
  el.cloudSave.addEventListener("click", cloudSave);
  el.cloudLoad.addEventListener("click", cloudLoad);

  window.addEventListener("pagehide", persistState);
  window.addEventListener("beforeunload", persistState);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
      persistState();
      return;
    }
    autoCloudCheckForRemoteUpdate();
  });
  window.addEventListener("focus", () => {
    autoCloudCheckForRemoteUpdate();
  });
  el.taskLargeFilter.addEventListener("change", () => {
    state.taskLargeFilterValue = el.taskLargeFilter.value;
    persistUiAndRender();
  });
  el.taskMidFilter.addEventListener("change", () => {
    state.taskMidFilterValue = el.taskMidFilter.value;
    persistUiAndRender();
  });
}

function renderAll() {
  renderVersionInfo();
  renderAddSections();
  renderGroupSelectors();
  renderManualSection();
  renderManualTaskOptions();
  renderTaskLargeFilter();
  renderTaskMidFilter();
  renderStatusFilter();
  renderTasks();
  renderActiveStatus();
  renderSummary();
  renderSummaryTabs();
}

function renderVersionInfo() {
  if (!el.appVersion) return;
  if (!el.appVersion.textContent || el.appVersion.textContent === "-" || el.appVersion.textContent === "算出中...") {
    el.appVersion.textContent = APP_VERSION_FALLBACK;
  }
}

async function initializeVersionInfo() {
  if (!el.appVersion) return;
  const dates = [];
  const htmlDate = parseDateSafe(document.lastModified);
  if (htmlDate) dates.push(htmlDate);

  const assetPaths = ["./app.js", "./style.css", "./manifest.webmanifest", "./service-worker.js", "./icon.svg"];
  const headerDates = await Promise.all(assetPaths.map((path) => fetchLastModified(path)));
  headerDates.forEach((d) => {
    if (d) dates.push(d);
  });

  if (!dates.length) {
    el.appVersion.textContent = APP_VERSION_FALLBACK;
    return;
  }

  const latestTime = Math.max(...dates.map((d) => d.getTime()));
  const latestDate = new Date(latestTime);
  el.appVersion.textContent = formatVersionDateTime(latestDate);
}

async function fetchLastModified(path) {
  try {
    const res = await fetch(path, {
      method: "GET",
      cache: "no-cache",
      headers: { "Cache-Control": "no-cache" },
    });
    if (!res.ok) return null;
    return parseDateSafe(res.headers.get("last-modified"));
  } catch {
    return null;
  }
}

function parseDateSafe(value) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function formatVersionDateTime(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day} ${hour}:${minute}`;
}

function renderAddSections() {
  const collapsed = Boolean(state.addSectionsCollapsed);
  el.addSectionsPanel.classList.toggle("collapsed", collapsed);
  el.addSectionsToggle.textContent = collapsed ? "開く" : "閉じる";
  el.addSectionsToggle.setAttribute("aria-expanded", String(!collapsed));
}

function renderGroupSelectors() {
  const topOptions = state.topGroups.map((g) => ({
    value: g.id,
    label: g.name,
    archived: Boolean(g.archived),
    entityType: "top",
    entityId: g.id,
  }));
  const largeOptions = state.largeGroups.map((g) => {
    const top = state.topGroups.find((t) => t.id === g.topGroupId);
    const topName = top ? top.name : "(不明な最上位グループ)";
    return {
      value: g.id,
      label: `${topName} > ${g.name}`,
      archived: Boolean(g.archived),
      entityType: "large",
      entityId: g.id,
    };
  });
  const taskParentOptions = getTaskParentOptions();

  const visibleTopOptions = resolveVisibleOptions(topOptions, "top");
  const visibleLargeOptions = resolveVisibleOptions(largeOptions, "large");
  const visibleTaskParentOptions = resolveVisibleOptions(taskParentOptions, "parent");

  state.uiSelections.largeTopId = ensureSelection(state.uiSelections.largeTopId, visibleTopOptions);
  state.uiSelections.midLargeId = ensureSelection(state.uiSelections.midLargeId, visibleLargeOptions);
  state.uiSelections.taskParentValue = ensureSelection(
    state.uiSelections.taskParentValue,
    visibleTaskParentOptions
  );

  renderDragDropdown(
    el.largeTopSelect,
    topOptions,
    state.uiSelections.largeTopId,
    "最上位グループを先に追加",
    state.archiveView.top,
    () => {
      state.archiveView.top = !state.archiveView.top;
      renderGroupSelectors();
    },
    (nextValue) => {
      state.uiSelections.largeTopId = nextValue;
      renderGroupSelectors();
    },
    (orderedValues) => {
      reorderByValues(state.topGroups, orderedValues);
      state.uiSelections.largeTopId = ensureSelection(
        state.uiSelections.largeTopId,
        filterArchivedOptions(topOptions, state.archiveView.top)
      );
      persistAndRender();
    },
    (entityType, entityId) => {
      toggleArchive(entityType, entityId);
      persistAndRender();
    }
  );

  renderDragDropdown(
    el.midLargeSelect,
    largeOptions,
    state.uiSelections.midLargeId,
    "大グループを先に追加",
    state.archiveView.large,
    () => {
      state.archiveView.large = !state.archiveView.large;
      renderGroupSelectors();
    },
    (nextValue) => {
      state.uiSelections.midLargeId = nextValue;
      renderGroupSelectors();
    },
    (orderedValues) => {
      reorderByValues(state.largeGroups, orderedValues);
      state.uiSelections.midLargeId = ensureSelection(
        state.uiSelections.midLargeId,
        filterArchivedOptions(largeOptions, state.archiveView.large)
      );
      persistAndRender();
    },
    (entityType, entityId) => {
      toggleArchive(entityType, entityId);
      persistAndRender();
    }
  );

  renderDragDropdown(
    el.taskMidSelect,
    taskParentOptions,
    state.uiSelections.taskParentValue,
    "大グループを先に追加",
    state.archiveView.parent,
    () => {
      state.archiveView.parent = !state.archiveView.parent;
      renderGroupSelectors();
    },
    (nextValue) => {
      state.uiSelections.taskParentValue = nextValue;
      renderGroupSelectors();
    },
    (orderedValues) => {
      state.taskParentOrder = orderedValues;
      persistAndRender();
    },
    (entityType, entityId) => {
      toggleArchive(entityType, entityId);
      persistAndRender();
    }
  );
}

function renderStatusFilter() {
  const selectedSet = new Set(state.taskFilterStatuses);
  const visibleCount = getFilteredTasks().length;
  const todayOnly = state.taskTodayFilterValue === "today";

  el.statusFilter.innerHTML = `
    <div class="status-filter-head">
      <p class="status-filter-label">表示フィルター</p>
      <p class="status-filter-count">${visibleCount}件表示中</p>
      <div class="today-filter-group">
        <button type="button" class="today-filter-btn ${todayOnly ? "" : "is-active"}" data-today-filter="all">全タスク</button>
        <button type="button" class="today-filter-btn ${todayOnly ? "is-active" : ""}" data-today-filter="today">本日のみ</button>
      </div>
      <div class="status-filter-actions">
        <button type="button" class="filter-action-btn" data-action="all">全選択</button>
        <button type="button" class="filter-action-btn" data-action="clear">すべて解除</button>
      </div>
    </div>
    <div class="status-filter-chips">
      ${STATUS_OPTIONS.map((status) => {
        const active = selectedSet.has(status) ? "is-active" : "";
        return `<button type="button" class="filter-chip ${active}" data-status="${status}">${status}</button>`;
      }).join("")}
    </div>
  `;

  el.statusFilter.querySelectorAll(".filter-chip").forEach((btn) => {
    btn.addEventListener("click", () => {
      const status = btn.dataset.status;
      const next = new Set(state.taskFilterStatuses);
      if (next.has(status)) {
        next.delete(status);
      } else {
        next.add(status);
      }
      state.taskFilterStatuses = STATUS_OPTIONS.filter((s) => next.has(s));
      persistUiAndRender();
    });
  });

  el.statusFilter.querySelectorAll(".filter-action-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const action = btn.dataset.action;
      if (action === "all") {
        state.taskFilterStatuses = [...STATUS_OPTIONS];
      } else {
        state.taskFilterStatuses = [];
      }
      persistUiAndRender();
    });
  });

  el.statusFilter.querySelectorAll(".today-filter-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.taskTodayFilterValue = btn.dataset.todayFilter === "today" ? "today" : "all";
      persistUiAndRender();
    });
  });
}

function renderTasks() {
  el.taskList.innerHTML = "";

  const activeTaskId = state.activeSession ? state.activeSession.taskId : null;
  const visibleTasks = getFilteredTasks();
  visibleTasks.sort((a, b) => {
    if (a.id === activeTaskId) return -1;
    if (b.id === activeTaskId) return 1;
    return 0;
  });

  visibleTasks.forEach((task) => {
    const node = el.taskTpl.content.firstElementChild.cloneNode(true);
    const path = resolveTaskPath(task.id);
    const isActive = Boolean(state.activeSession && state.activeSession.taskId === task.id);
    node.dataset.taskId = task.id;
    node.draggable = !isActive;

    node.querySelector(".task-title").textContent = task.name;
    node.querySelector(".task-path").textContent = `${path.topName} > ${path.largeName} > ${path.midName} > ${task.name}`;
    node.querySelector(".task-status").textContent = `状態: ${normalizeStatus(task.status)}${isActive ? " (計測中)" : ""}`;
    node.querySelector(".task-today").textContent = `本日対象: ${task.isTodayTask ? "はい" : "いいえ"}`;
    node.querySelector(".task-repeat").textContent = `繰り返し: ${formatRecurrenceLabel(task.recurrence)}`;
    node.querySelector(".task-tags").textContent = `タグ: ${(task.tags || []).join(", ") || "なし"}`;
    node.classList.toggle("is-active", isActive);
    node.classList.toggle("is-completed", normalizeStatus(task.status) === "完了");
    node.classList.toggle("is-withdrawn", normalizeStatus(task.status) === "取り下げ");
    node.classList.toggle("is-today-task", Boolean(task.isTodayTask));

    const statusSelect = node.querySelector(".task-status-select");
    statusSelect.innerHTML = STATUS_OPTIONS.map(
      (status) => `<option value="${status}">${status}</option>`
    ).join("");
    statusSelect.value = normalizeStatus(task.status);
    statusSelect.addEventListener("change", () => {
      task.status = normalizeStatus(statusSelect.value);
      task.updatedAt = Date.now();
      persistAndRender();
    });

    const repeatSelect = node.querySelector(".task-repeat-select");
    repeatSelect.innerHTML = REPEAT_OPTIONS.map(
      (repeat) => `<option value="${repeat}">${formatRecurrenceLabel(repeat)}</option>`
    ).join("");
    repeatSelect.value = normalizeRecurrence(task.recurrence);
    repeatSelect.addEventListener("change", () => {
      task.recurrence = normalizeRecurrence(repeatSelect.value);
      task.recurrenceResetKey = task.recurrence === "none" ? "" : getRecurrencePeriodKey(task.recurrence);
      task.updatedAt = Date.now();
      persistAndRender();
    });

    const startBtn = node.querySelector(".start-btn");
    const stopBtn = node.querySelector(".stop-btn");
    const todayBtn = node.querySelector(".today-btn");
    const renameBtn = node.querySelector(".rename-btn");
    const deleteBtn = node.querySelector(".delete-btn");
    startBtn.addEventListener("click", () => startTask(task.id));
    stopBtn.addEventListener("click", () => stopTask(task.id));
    todayBtn.textContent = task.isTodayTask ? "本日対象から外す" : "本日対象にする";
    todayBtn.classList.toggle("is-active", Boolean(task.isTodayTask));
    todayBtn.addEventListener("click", () => {
      task.isTodayTask = !task.isTodayTask;
      task.updatedAt = Date.now();
      persistAndRender();
    });
    renameBtn.addEventListener("click", () => renameTask(task.id));
    deleteBtn.addEventListener("click", () => deleteTask(task.id));

    if (isActive) {
      startBtn.disabled = true;
      stopBtn.disabled = false;
      node.querySelector(".task-drag-handle").style.opacity = "0.4";
      node.querySelector(".task-drag-handle").style.cursor = "not-allowed";
    } else {
      startBtn.disabled = Boolean(state.activeSession);
      stopBtn.disabled = true;
    }

    node.addEventListener("dragstart", () => {
      draggingTaskId = task.id;
      node.classList.add("is-dragging");
    });
    node.addEventListener("dragend", () => {
      draggingTaskId = "";
      node.classList.remove("is-dragging");
      document.querySelectorAll(".task-row.drag-over-task").forEach((row) => {
        row.classList.remove("drag-over-task");
      });
    });
    node.addEventListener("dragover", (e) => {
      e.preventDefault();
      if (!draggingTaskId || draggingTaskId === task.id) return;
      node.classList.add("drag-over-task");
    });
    node.addEventListener("dragleave", () => {
      node.classList.remove("drag-over-task");
    });
    node.addEventListener("drop", (e) => {
      e.preventDefault();
      node.classList.remove("drag-over-task");
      if (!draggingTaskId || draggingTaskId === task.id) return;
      const moved = reorderTaskByIds(draggingTaskId, task.id);
      if (moved) {
        persistAndRender();
      }
    });

    el.taskList.appendChild(node);
  });

  if (!visibleTasks.length) {
    el.taskList.innerHTML = "<p>表示対象のタスクがありません</p>";
  }
}

function renderTaskMidFilter() {
  const options = getTaskMidFilterOptions();
  state.taskMidFilterValue = ensureSelection(state.taskMidFilterValue, options);
  el.taskMidFilter.innerHTML = options
    .map((opt) => `<option value="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</option>`)
    .join("");
  el.taskMidFilter.value = state.taskMidFilterValue;
}

function renderTaskLargeFilter() {
  const options = getTaskLargeFilterOptions();
  state.taskLargeFilterValue = ensureSelection(state.taskLargeFilterValue, options);
  el.taskLargeFilter.innerHTML = options
    .map((opt) => `<option value="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</option>`)
    .join("");
  el.taskLargeFilter.value = state.taskLargeFilterValue;
}

function reorderTaskByIds(fromId, toId) {
  const fromIndex = state.tasks.findIndex((t) => t.id === fromId);
  const toIndex = state.tasks.findIndex((t) => t.id === toId);
  if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) return false;
  const [moved] = state.tasks.splice(fromIndex, 1);
  state.tasks.splice(toIndex, 0, moved);
  return true;
}

function renderActiveStatus() {
  if (!state.activeSession) {
    el.activeStatus.textContent = "稼働中タスク: なし";
    return;
  }

  const task = state.tasks.find((t) => t.id === state.activeSession.taskId);
  if (!task) {
    el.activeStatus.textContent = "稼働中タスク: 不明";
    return;
  }

  const elapsedMs = Date.now() - state.activeSession.startAt;
  el.activeStatus.textContent = `稼働中タスク: ${task.name} (${formatDuration(elapsedMs)})`;
}

function renderSummary() {
  const range = el.rangeSelect.value;
  const now = Date.now();
  const offset = state.summaryOffsets[range] || 0;
  const bounds = getRangeBounds(range, now, offset);
  const start = bounds.start;
  const end = Math.min(bounds.end, now);

  const sessions = materializeSessions(now).filter(
    (s) => s.endAt > start && s.startAt < end
  );

  const taskMap = new Map();
  const midMap = new Map();
  const largeMap = new Map();
  const topMap = new Map();
  const tagMap = new Map();
  let totalMs = 0;

  sessions.forEach((session) => {
    const overlapMs = overlap(session.startAt, session.endAt, start, end);
    if (overlapMs <= 0) return;
    totalMs += overlapMs;

    addDuration(taskMap, session.taskLabel, overlapMs);
    addDuration(midMap, session.midName, overlapMs);
    addDuration(largeMap, session.largeName, overlapMs);
    addDuration(topMap, session.topName, overlapMs);

    session.allTags.forEach((tag) => addDuration(tagMap, tag, overlapMs));
  });

  el.summaryPeriod.textContent = `対象: ${formatRangeLabel(range, bounds)}`;
  el.summaryTotal.textContent = `期間合計: ${formatDuration(totalMs)}`;
  el.summaryNext.disabled = offset >= 0;
  el.summaryReset.disabled = offset === 0;

  renderList(el.summaryTask, taskMap, totalMs, "task");
  renderList(el.summaryMid, midMap, totalMs, "mid");
  renderList(el.summaryLarge, largeMap, totalMs, "large");
  renderList(el.summaryTop, topMap, totalMs, "top");
  renderList(el.summaryTags, tagMap, totalMs, "tags");
}

function renderSummaryTabs() {
  const activeTab = normalizeSummaryTab(state.summaryTab);
  el.summaryTabs.querySelectorAll(".summary-tab-btn").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.tab === activeTab);
  });
  document.querySelectorAll(".summary-tab-panel").forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.panel === activeTab);
  });
}

function renderList(container, map, totalMs, sectionKey) {
  const items = [...map.entries()].sort((a, b) => b[1] - a[1]);
  if (!items.length) {
    container.innerHTML = "<li>データなし</li>";
    return;
  }

  const maxMs = items[0][1] || 1;
  const isExpanded = Boolean(state.summaryExpanded[sectionKey]);
  const visibleLimit = isExpanded ? items.length : 8;
  const visibleItems = items.slice(0, visibleLimit);
  container.innerHTML = visibleItems
    .map(([name, ms]) => {
      const rate = totalMs > 0 ? (ms / totalMs) * 100 : 0;
      const barRate = (ms / maxMs) * 100;
      return `<li>
        <div class="summary-item-head">
          <span class="summary-item-name">${escapeHtml(name)}</span>
          <span class="summary-item-value">${formatDuration(ms)}</span>
        </div>
        <div class="summary-bar"><div class="summary-bar-fill" style="width:${barRate.toFixed(1)}%"></div></div>
        <div class="summary-item-rate">${rate.toFixed(1)}%</div>
      </li>`;
    })
    .join("");

  if (items.length > 8) {
    const remaining = items.length - 8;
    const label = isExpanded ? "折りたたむ" : `ほか ${remaining} 件を表示`;
    container.innerHTML += `<li class="summary-more-row"><button type="button" class="summary-more-btn" data-summary-expand="${sectionKey}">${label}</button></li>`;
  }

  const expandBtn = container.querySelector(".summary-more-btn");
  if (expandBtn) {
    expandBtn.addEventListener("click", () => {
      state.summaryExpanded[sectionKey] = !Boolean(state.summaryExpanded[sectionKey]);
      persistUiAndRender();
    });
  }
}

function startTask(taskId) {
  if (state.activeSession) {
    alert("同時に開始できるタスクは1つです。稼働中タスクを終了してください。");
    return;
  }
  const task = state.tasks.find((t) => t.id === taskId);
  if (task) {
    task.status = "着手";
    task.updatedAt = Date.now();
  }
  state.activeSession = { taskId, startAt: Date.now() };
  persistAndRender();
}

function stopTask(taskId) {
  if (!state.activeSession || state.activeSession.taskId !== taskId) return;

  const startAt = state.activeSession.startAt;
  const endAt = Date.now();
  const snapshot = snapshotTask(taskId);
  state.sessions.push({
    id: uid(),
    taskId,
    startAt,
    endAt,
    snapshot,
    updatedAt: Date.now(),
  });
  state.activeSession = null;
  persistAndRender("now");
}

function deleteTask(taskId) {
  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;
  if (state.activeSession && state.activeSession.taskId === taskId) {
    alert("計測中のタスクは削除できません。先に終了してください。");
    return;
  }
  const ok = confirm(`タスク「${task.name}」を削除しますか？`);
  if (!ok) return;
  state.tasks = state.tasks.filter((t) => t.id !== taskId);
  if (!state.deletedTaskIds.includes(taskId)) {
    state.deletedTaskIds.push(taskId);
  }
  persistAndRender();
}

function renameTask(taskId) {
  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;
  const nextName = prompt("新しいタスク名を入力してください", task.name);
  if (nextName === null) return;
  const trimmed = nextName.trim();
  if (!trimmed) {
    alert("タスク名は空にできません。");
    return;
  }
  if (trimmed === task.name) return;
  task.name = trimmed;
  task.updatedAt = Date.now();
  persistAndRender();
}

function addManualSession(taskId, startAt, endAt) {
  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) {
    alert("対象タスクが見つかりません。");
    return;
  }
  const snapshot = snapshotTask(taskId);
  state.sessions.push({
    id: uid(),
    taskId,
    startAt,
    endAt,
    snapshot,
    updatedAt: Date.now(),
  });
}

function exportBackup() {
  const payload = {
    version: 1,
    exportedAt: new Date().toISOString(),
    data: JSON.parse(JSON.stringify(state)),
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `work-tracker-backup-${stamp}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function importBackup(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const raw = parsed && typeof parsed === "object" && parsed.data ? parsed.data : parsed;
    const migrated = migrateState(raw);

    const ok = confirm("現在のデータを上書きして復元しますか？");
    if (!ok) return;

    replaceState(migrated);
    persistAndRender();
  } catch {
    alert("JSON読込に失敗しました。バックアップファイルを確認してください。");
  } finally {
    el.backupFile.value = "";
  }
}

function replaceState(next) {
  state.topGroups = next.topGroups;
  state.largeGroups = next.largeGroups;
  state.midGroups = next.midGroups;
  state.tasks = next.tasks;
  state.sessions = next.sessions;
  state.deletedTaskIds = next.deletedTaskIds;
  state.deletedSessionIds = next.deletedSessionIds;
  state.activeSession = next.activeSession;
  state.taskFilterStatuses = next.taskFilterStatuses;
  state.summaryTab = next.summaryTab;
  state.summaryOffsets = next.summaryOffsets;
  state.summaryExpanded = next.summaryExpanded;
  state.taskParentOrder = next.taskParentOrder;
  state.uiSelections = next.uiSelections;
  state.archiveView = next.archiveView;
  state.taskLargeFilterValue = next.taskLargeFilterValue;
  state.taskMidFilterValue = next.taskMidFilterValue;
  state.taskTodayFilterValue = next.taskTodayFilterValue;
  state.todayTaskDateKey = next.todayTaskDateKey;
  state.manualCollapsed = next.manualCollapsed;
  state.timelineView = next.timelineView;
  state.addSectionsCollapsed = next.addSectionsCollapsed;
  state.lastDataChangeAt = next.lastDataChangeAt;
}

async function cloudSave() {
  if (isCloudSyncBusy) return;
  const endpoint = getCloudEndpoint();
  if (!endpoint) return;

  isCloudSyncBusy = true;
  setCloudBusy(true);
  try {
    await cloudSaveRequest(endpoint);
    alert("クラウド保存が完了しました。");
  } catch (error) {
    alert(error.message || "クラウド保存に失敗しました。");
  } finally {
    isCloudSyncBusy = false;
    setCloudBusy(false);
  }
}

async function cloudLoad() {
  if (isCloudSyncBusy) return;
  const endpoint = getCloudEndpoint();
  if (!endpoint) return;

  isCloudSyncBusy = true;
  setCloudBusy(true);
  try {
    const parsed = await cloudLoadRequest(endpoint);
    if (!parsed.data) {
      alert("クラウドに復元データがありません。");
      return;
    }
    const migrated = migrateState(parsed.data);
    lastSeenCloudSavedAt = Math.max(
      lastSeenCloudSavedAt,
      Number.isFinite(parsed.savedAt) ? parsed.savedAt : resolveStateUpdatedAt(migrated)
    );
    const ok = confirm("クラウドデータで現在データを上書きしますか？");
    if (!ok) return;
    replaceState(migrated);
    persistUiAndRender();
  } catch (error) {
    alert(error.message || "クラウド読込に失敗しました。");
  } finally {
    isCloudSyncBusy = false;
    setCloudBusy(false);
  }
}

async function cloudSaveRequest(endpoint) {
  const mergedData = await mergeStateWithCloud(endpoint, state);
  const savedAt = resolveStateUpdatedAt(mergedData);
  const payload = {
    action: "save",
    sheetId: SPREADSHEET_ID,
    savedAt,
    data: mergedData,
  };
  const res = await fetch(endpoint, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  const text = await res.text();
  const parsed = parseCloudResponse(text);
  if (!res.ok || !parsed.ok) {
    throw new Error(parsed.message || "クラウド保存に失敗しました。");
  }
  lastSeenCloudSavedAt = Math.max(lastSeenCloudSavedAt, savedAt);
  replaceState(migrateState(mergedData));
  persistState();
  return parsed;
}

async function cloudLoadRequest(endpoint) {
  const url = new URL(endpoint);
  url.searchParams.set("action", "load");
  url.searchParams.set("sheetId", SPREADSHEET_ID);
  const res = await fetch(url.toString(), { method: "GET" });
  const text = await res.text();
  const parsed = parseCloudResponse(text);
  if (!res.ok || !parsed.ok) {
    throw new Error(parsed.message || "クラウド読込に失敗しました。");
  }
  return parsed;
}

async function cloudLatestSavedAtRequest(endpoint) {
  const url = new URL(endpoint);
  url.searchParams.set("action", "latest");
  url.searchParams.set("sheetId", SPREADSHEET_ID);
  const res = await fetch(url.toString(), { method: "GET" });
  const text = await res.text();
  const parsed = parseCloudResponse(text);
  if (!res.ok || !parsed.ok) {
    throw new Error(parsed.message || "クラウド更新確認に失敗しました。");
  }
  return Number.isFinite(parsed.savedAt) ? parsed.savedAt : 0;
}

function parseCloudResponse(text) {
  try {
    return JSON.parse(text);
  } catch {
    return { ok: false, message: "レスポンスがJSON形式ではありません。" };
  }
}

function getCloudEndpoint() {
  const endpoint = el.cloudEndpoint.value.trim();
  if (!endpoint) {
    alert("Apps Script WebアプリURLを入力してください。");
    return "";
  }
  localStorage.setItem(CLOUD_ENDPOINT_KEY, endpoint);
  return endpoint;
}

function setCloudBusy(isBusy) {
  el.cloudSave.disabled = isBusy;
  el.cloudLoad.disabled = isBusy;
}

function startCloudAutoSync() {
  autoCloudLoadOnStartup();
  setInterval(() => {
    autoCloudSave();
  }, AUTO_CLOUD_SAVE_INTERVAL_MS);
  setInterval(() => {
    autoCloudCheckForRemoteUpdate();
  }, AUTO_CLOUD_LATEST_CHECK_INTERVAL_MS);
  setInterval(() => {
    autoCloudLoadPeriodic();
  }, AUTO_CLOUD_LOAD_INTERVAL_MS);
}

async function autoCloudLoadOnStartup() {
  const endpoint = el.cloudEndpoint.value.trim();
  if (!endpoint || !isStateEmpty(state) || isCloudSyncBusy) return;

  isCloudSyncBusy = true;
  setCloudBusy(true);
  try {
    const parsed = await cloudLoadRequest(endpoint);
    if (!parsed.data) return;
    const migrated = migrateState(parsed.data);
    lastSeenCloudSavedAt = Math.max(
      lastSeenCloudSavedAt,
      Number.isFinite(parsed.savedAt) ? parsed.savedAt : resolveStateUpdatedAt(migrated)
    );
    replaceState(migrated);
    persistUiAndRender();
  } catch {
    // Silent failure for background startup sync.
  } finally {
    isCloudSyncBusy = false;
    setCloudBusy(false);
  }
}

async function autoCloudCheckForRemoteUpdate() {
  const endpoint = el.cloudEndpoint.value.trim();
  if (!endpoint || isCloudSyncBusy) return;
  if (state.activeSession) return;
  try {
    const latestSavedAt = await cloudLatestSavedAtRequest(endpoint);
    if (!latestSavedAt || latestSavedAt <= lastSeenCloudSavedAt) return;
    await autoCloudLoadPeriodic();
  } catch {
    // Silent failure for background update checks.
  }
}

async function autoCloudLoadPeriodic() {
  const endpoint = el.cloudEndpoint.value.trim();
  if (!endpoint || isCloudSyncBusy) return;

  isCloudSyncBusy = true;
  setCloudBusy(true);
  try {
    const parsed = await cloudLoadRequest(endpoint);
    if (!parsed.data) return;
    lastSeenCloudSavedAt = Math.max(
      lastSeenCloudSavedAt,
      Number.isFinite(parsed.savedAt) ? parsed.savedAt : 0
    );
    const remote = migrateState(parsed.data);
    const local = migrateState(state);
    const merged = mergeStates(remote, local);
    const localStamp = resolveStateUpdatedAt(local);
    const mergedStamp = resolveStateUpdatedAt(merged, parsed.savedAt);
    if (mergedStamp < localStamp) return;
    if (state.activeSession && !merged.activeSession) {
      merged.activeSession = state.activeSession;
    }
    const localSerialized = JSON.stringify(local);
    const mergedSerialized = JSON.stringify(merged);
    if (localSerialized === mergedSerialized) return;
    replaceState(merged);
    persistUiAndRender();
  } catch {
    // Silent failure for background periodic sync.
  } finally {
    isCloudSyncBusy = false;
    setCloudBusy(false);
  }
}

async function autoCloudSave() {
  const endpoint = el.cloudEndpoint.value.trim();
  if (!endpoint || isCloudSyncBusy) return;

  isCloudSyncBusy = true;
  setCloudBusy(true);
  try {
    await cloudSaveRequest(endpoint);
  } catch {
    // Silent failure for background autosave.
  } finally {
    isCloudSyncBusy = false;
    setCloudBusy(false);
  }
}

function snapshotTask(taskId) {
  const path = resolveTaskPath(taskId);
  const task = state.tasks.find((t) => t.id === taskId);
  const customTags = task?.tags || [];

  return {
    taskName: task?.name || "(削除済みタスク)",
    taskLabel: `${path.topName} > ${path.largeName} > ${path.midName} > ${task?.name || "(削除済みタスク)"}`,
    midName: path.midName,
    largeName: path.largeName,
    topName: path.topName,
    customTags,
  };
}

function materializeSessions(now) {
  const history = state.sessions.map((s) => {
    const snap = s.snapshot || {};
    const allTags = buildAllTags(snap);
    return {
      sessionId: s.id,
      isActive: false,
      startAt: s.startAt,
      endAt: s.endAt,
      taskName: snap.taskName || "(削除済みタスク)",
      taskLabel:
        snap.taskLabel ||
        `${snap.topName || "(削除済み最上位グループ)"} > ${snap.largeName || "(削除済み大グループ)"} > ${snap.midName || "(削除済み中グループ)"} > ${snap.taskName || "(削除済みタスク)"}`,
      midName: snap.midName || "(削除済み中グループ)",
      largeName: snap.largeName || "(削除済み大グループ)",
      topName: snap.topName || "(削除済み最上位グループ)",
      customTags: Array.isArray(snap.customTags) ? snap.customTags : [],
      allTags,
    };
  });

  if (state.activeSession) {
    const snap = snapshotTask(state.activeSession.taskId);
    history.push({
      sessionId: "",
      isActive: true,
      startAt: state.activeSession.startAt,
      endAt: now,
      taskName: snap.taskName,
      taskLabel: snap.taskLabel,
      midName: snap.midName,
      largeName: snap.largeName,
      topName: snap.topName,
      customTags: Array.isArray(snap.customTags) ? snap.customTags : [],
      allTags: buildAllTags(snap),
    });
  }

  return history;
}

function buildAllTags(snap) {
  const tags = new Set();
  tags.add(`最上位:${snap.topName || "不明"}`);
  tags.add(`大:${snap.largeName || "不明"}`);
  tags.add(`中:${snap.midName || "不明"}`);
  tags.add(`タスク:${snap.taskName || "不明"}`);
  (snap.customTags || []).forEach((t) => tags.add(`タグ:${t}`));
  return [...tags];
}

function resolveTaskPath(taskId) {
  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) {
    return {
      midName: "(削除済み中グループ)",
      largeName: "(削除済み大グループ)",
      topName: "(削除済み最上位グループ)",
    };
  }

  let mid = null;
  let large = null;
  if (task.parentType === "large") {
    large = state.largeGroups.find((g) => g.id === task.largeGroupId) || null;
  } else {
    mid = state.midGroups.find((m) => m.id === task.midGroupId) || null;
    large = mid ? state.largeGroups.find((g) => g.id === mid.largeGroupId) || null : null;
  }
  const top = large ? state.topGroups.find((tg) => tg.id === large.topGroupId) : null;

  return {
    midName: task.parentType === "large" ? "(中グループなし)" : mid?.name || "(削除済み中グループ)",
    largeName: large?.name || "(削除済み大グループ)",
    topName: top?.name || "(削除済み最上位グループ)",
  };
}

function renderManualTaskOptions() {
  const options = state.tasks
    .filter((task) => {
      const status = normalizeStatus(task.status);
      return status !== "完了" && status !== "取り下げ";
    })
    .map((task) => {
      const path = resolveTaskPath(task.id);
      const label = `${path.topName} > ${path.largeName} > ${path.midName} > ${task.name}`;
      return `<option value="${task.id}">${escapeHtml(label)}</option>`;
    })
    .join("");
  el.manualTaskSelect.innerHTML = options || `<option value="">タスクを先に追加</option>`;
}

function getFilteredTasks() {
  const visibleStatuses = new Set(state.taskFilterStatuses);
  const selectedLarge = state.taskLargeFilterValue || "all";
  const selectedMid = state.taskMidFilterValue || "all";
  const todayOnly = state.taskTodayFilterValue === "today";
  const activeTaskId = state.activeSession ? state.activeSession.taskId : "";
  return state.tasks.filter((task) => {
    if (task.id === activeTaskId) return true;
    if (!visibleStatuses.has(normalizeStatus(task.status))) return false;
    if (todayOnly && !task.isTodayTask) return false;
    if (selectedLarge !== "all" && getTaskLargeFilterKey(task) !== selectedLarge) return false;
    if (selectedMid === "all") return true;
    return getTaskMidFilterKey(task) === selectedMid;
  });
}

function getTaskLargeFilterOptions() {
  const map = new Map();
  map.set("all", { value: "all", label: "すべて" });
  state.tasks.forEach((task) => {
    const key = getTaskLargeFilterKey(task);
    if (!map.has(key)) {
      map.set(key, { value: key, label: getTaskLargeFilterLabel(task) });
    }
  });
  return [...map.values()];
}

function getTaskLargeFilterKey(task) {
  const path = resolveTaskPath(task.id);
  if (path.largeName === "(削除済み大グループ)") return "large:unknown";
  if (task.parentType === "large" && task.largeGroupId) return `large:${task.largeGroupId}`;
  if (task.parentType === "mid" && task.midGroupId) {
    const mid = state.midGroups.find((m) => m.id === task.midGroupId);
    if (mid && mid.largeGroupId) return `large:${mid.largeGroupId}`;
  }
  return "large:unknown";
}

function getTaskLargeFilterLabel(task) {
  const path = resolveTaskPath(task.id);
  if (path.largeName === "(削除済み大グループ)") return "(削除済み大グループ)";
  return `${path.topName} > ${path.largeName}`;
}

function getTaskMidFilterOptions() {
  const map = new Map();
  map.set("all", { value: "all", label: "すべて" });
  state.tasks.forEach((task) => {
    const key = getTaskMidFilterKey(task);
    if (!map.has(key)) {
      map.set(key, { value: key, label: getTaskMidFilterLabel(task) });
    }
  });
  return [...map.values()];
}

function getTaskMidFilterKey(task) {
  if (task.parentType === "large") return "mid:none";
  if (task.midGroupId) return `mid:${task.midGroupId}`;
  return "mid:unknown";
}

function getTaskMidFilterLabel(task) {
  const path = resolveTaskPath(task.id);
  if (task.parentType === "large") return "(中グループなし)";
  if (task.midGroupId) return `${path.largeName} > ${path.midName}`;
  return "(削除済み中グループ)";
}

function renderManualSection() {
  const collapsed = Boolean(state.manualCollapsed);
  el.manualPanel.classList.toggle("collapsed", collapsed);
  el.manualToggle.textContent = collapsed ? "開く" : "閉じる";
  el.manualToggle.setAttribute("aria-expanded", String(!collapsed));
}

function openTimelineModal() {
  if (!el.timelineDayDate.value) {
    el.timelineDayDate.value = formatDateInput(new Date());
  }
  el.timelineModal.classList.remove("hidden");
  renderTimelineModal();
}

function closeTimelineModal() {
  el.timelineModal.classList.add("hidden");
}

function renderTimelineModal() {
  const sessions = materializeSessions(Date.now());
  const dayText = el.timelineDayDate.value || formatDateInput(new Date());
  el.timelineDayDate.value = dayText;
  const keyword = el.timelineKeyword.value.trim().toLowerCase();

  const dayStart = new Date(`${dayText}T00:00:00`).getTime();
  const dayEnd = dayStart + 24 * 60 * 60 * 1000;

  const filtered = sessions
    .filter((s) => s.endAt > dayStart && s.startAt < dayEnd)
    .filter((s) => {
      if (!keyword) return true;
      const haystack = `${s.taskLabel} ${s.taskName} ${s.midName} ${s.largeName} ${s.topName}`.toLowerCase();
      return haystack.includes(keyword);
    })
    .sort((a, b) => a.startAt - b.startAt);

  const view = normalizeTimelineView(state.timelineView);
  el.timelineViewToggle.querySelectorAll("button[data-view]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.view === view);
  });

  const totalMs = filtered.reduce(
    (sum, s) => sum + overlap(s.startAt, s.endAt, dayStart, dayEnd),
    0
  );
  el.timelineMeta.textContent = `${formatTimelineDayLabel(dayStart)} / ${filtered.length}件 / 合計 ${formatDuration(totalMs)}`;

  if (view === "list") {
    renderTimelineListView(filtered);
    return;
  }

  const hourRows = Array.from({ length: 24 }, (_, h) => {
    const label = `${String(h).padStart(2, "0")}:00`;
    return `<div class="timeline-hour"><span>${label}</span></div>`;
  }).join("");

  const positionedEvents = layoutTimelineEvents(
    filtered.map((s, index) => {
      const startMin = clamp(
        Math.floor((s.startAt - dayStart) / 60000),
        0,
        24 * 60
      );
      const endMin = clamp(
        Math.ceil((s.endAt - dayStart) / 60000),
        0,
        24 * 60
      );
      const durationMin = Math.max(TIMELINE_MIN_EVENT_HEIGHT, endMin - startMin);
      const color = timelineColorByText(s.taskLabel || s.taskName || String(index));
      return {
        sessionId: s.sessionId,
        isActive: Boolean(s.isActive),
        taskName: s.taskName,
        taskLabel: s.taskLabel,
        customTags: Array.isArray(s.customTags) ? s.customTags : [],
        startAt: s.startAt,
        endAt: s.endAt,
        startMin,
        endMin,
        visualEndMin: startMin + durationMin,
        durationMin,
        color,
      };
    })
  );

  const events = positionedEvents
    .map(
      (e) => {
        const laneWidth = 100 / e.columns;
        const left = laneWidth * e.lane;
        const tight = e.durationMin < 30 ? "is-tight" : "";
        const dense = e.columns >= 4 ? "is-dense" : "";
        const detailTitle = encodeURIComponent(e.taskLabel);
        const detailRange = encodeURIComponent(
          `${formatTimelineDateTime(e.startAt)}〜${formatTimelineDateTime(e.endAt)}：${formatDuration(e.endAt - e.startAt)}`
        );
        const detailTags = encodeURIComponent((e.customTags || []).join(", "));
        const sessionId = e.sessionId || "";
        const isActive = e.isActive ? "1" : "0";
        const deleteBtn =
          sessionId && !e.isActive
            ? `<button type="button" class="timeline-event-delete-btn" data-session-id="${escapeHtml(sessionId)}" title="この記録を削除">×</button>`
            : "";
        return `<article class="timeline-event ${tight} ${dense}" data-session-id="${escapeHtml(sessionId)}" data-is-active="${isActive}" data-detail-title="${detailTitle}" data-detail-range="${detailRange}" data-detail-tags="${detailTags}" data-detail-color="${e.color}" style="top:${e.startMin}px;height:${e.durationMin}px;left:calc(${left}% + 10px);width:calc(${laneWidth}% - 14px);border-left-color:${e.color};background:${e.color}22;">
          ${deleteBtn}
          <p class="timeline-event-time">${formatTimelineTime(e.startAt)} - ${formatTimelineTime(e.endAt)}</p>
          <p class="timeline-event-label">${escapeHtml(e.taskLabel)}</p>
        </article>`;
      }
    )
    .join("");

  el.timelineList.innerHTML = `
    <section class="timeline-day-calendar">
      <div class="timeline-hours">${hourRows}</div>
      <div class="timeline-track">
        ${events}
      </div>
    </section>
  `;
  bindTimelineHoverDetail();
}

function renderTimelineListView(items) {
  if (!items.length) {
    el.timelineList.innerHTML = `<p class="timeline-empty">表示データがありません</p>`;
    return;
  }
  el.timelineList.innerHTML = `<section class="timeline-legacy-list">
    ${items
      .map((s) => {
        const tags = (s.customTags || []).join(", ");
        const tagLine = tags ? `<p class="timeline-row-tags">タグ: ${escapeHtml(tags)}</p>` : "";
        const editBtn = s.isActive
          ? ""
          : `<button type="button" class="timeline-edit-btn" data-session-id="${escapeHtml(s.sessionId)}">時間修正</button>`;
        const deleteBtn = s.isActive
          ? ""
          : `<button type="button" class="timeline-delete-btn" data-session-id="${escapeHtml(s.sessionId)}">削除</button>`;
        return `<article class="timeline-legacy-row">
          <div class="timeline-legacy-main">
            <p class="timeline-legacy-time">${formatTimelineDateTime(s.startAt)}〜${formatTimelineDateTime(s.endAt)}：${formatDuration(
              s.endAt - s.startAt
            )}</p>
            <p class="timeline-legacy-label">${escapeHtml(s.taskLabel)}</p>
            ${tagLine}
          </div>
          <div class="timeline-legacy-actions">${editBtn}${deleteBtn}</div>
        </article>`;
      })
      .join("")}
  </section>`;
  el.timelineList.querySelectorAll(".timeline-edit-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      editSessionTime(btn.dataset.sessionId);
    });
  });
  el.timelineList.querySelectorAll(".timeline-delete-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      deleteSession(btn.dataset.sessionId);
    });
  });
  el.timelineHoverDetail.classList.add("hidden");
  el.timelineHoverDetail.classList.remove("is-visible");
}

function shiftTimelineDay(diffDays) {
  const base = el.timelineDayDate.value
    ? new Date(`${el.timelineDayDate.value}T00:00:00`)
    : new Date();
  base.setDate(base.getDate() + diffDays);
  el.timelineDayDate.value = formatDateInput(base);
  renderTimelineModal();
}

function formatTimelineDayLabel(ms) {
  const d = new Date(ms);
  const week = ["日", "月", "火", "水", "木", "金", "土"][d.getDay()];
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}(${week})`;
}

function timelineColorByText(text) {
  const palette = ["#0f766e", "#0ea5e9", "#1d4ed8", "#7c3aed", "#db2777", "#ea580c", "#16a34a"];
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = (hash << 5) - hash + text.charCodeAt(i);
    hash |= 0;
  }
  return palette[Math.abs(hash) % palette.length];
}

function layoutTimelineEvents(events) {
  const sorted = [...events].sort((a, b) => a.startMin - b.startMin || a.endMin - b.endMin);
  const active = [];

  sorted.forEach((event) => {
    for (let i = active.length - 1; i >= 0; i -= 1) {
      if (active[i].visualEndMin <= event.startMin) {
        active.splice(i, 1);
      }
    }
    const used = new Set(active.map((item) => item.lane));
    let lane = 0;
    while (used.has(lane)) lane += 1;
    event.lane = lane;
    active.push({ visualEndMin: event.visualEndMin, lane });
  });

  let cluster = [];
  let clusterEnd = -1;
  for (let i = 0; i < sorted.length; i += 1) {
    const event = sorted[i];
    if (!cluster.length || event.startMin < clusterEnd) {
      cluster.push(event);
      clusterEnd = Math.max(clusterEnd, event.visualEndMin);
      continue;
    }
    applyTimelineClusterColumns(cluster);
    cluster = [event];
    clusterEnd = event.visualEndMin;
  }
  applyTimelineClusterColumns(cluster);

  return sorted;
}

function applyTimelineClusterColumns(cluster) {
  if (!cluster.length) return;
  const points = [];
  cluster.forEach((event) => {
    points.push({ time: event.startMin, delta: 1 });
    points.push({ time: event.visualEndMin, delta: -1 });
  });
  points.sort((a, b) => {
    if (a.time === b.time) return a.delta - b.delta;
    return a.time - b.time;
  });
  let current = 0;
  let max = 1;
  points.forEach((p) => {
    current += p.delta;
    if (current > max) max = current;
  });
  cluster.forEach((event) => {
    event.columns = max;
  });
}

function bindTimelineHoverDetail() {
  el.timelineHoverDetail.classList.add("hidden");
  el.timelineHoverDetail.classList.remove("is-visible");
  el.timelineHoverDetail.innerHTML = "";

  const events = el.timelineList.querySelectorAll(".timeline-event");
  events.forEach((eventNode) => {
    const show = () => {
      const title = decodeURIComponent(eventNode.dataset.detailTitle || "");
      const range = decodeURIComponent(eventNode.dataset.detailRange || "");
      const tags = decodeURIComponent(eventNode.dataset.detailTags || "");
      const color = eventNode.dataset.detailColor || "#0f766e";
      el.timelineHoverDetail.innerHTML = `
        <div class="timeline-hover-head" style="--detail-color:${escapeHtml(color)}">
          <span class="timeline-hover-dot"></span>
          <p class="timeline-hover-title">${escapeHtml(title)}</p>
        </div>
        <p class="timeline-hover-range">${escapeHtml(range)}</p>
        ${tags ? `<p class="timeline-hover-tags">タグ: ${escapeHtml(tags)}</p>` : ""}
      `;
      el.timelineHoverDetail.classList.remove("hidden");
      requestAnimationFrame(() => {
        el.timelineHoverDetail.classList.add("is-visible");
      });
    };
    const hide = () => {
      el.timelineHoverDetail.classList.remove("is-visible");
      setTimeout(() => {
        if (!el.timelineHoverDetail.classList.contains("is-visible")) {
          el.timelineHoverDetail.classList.add("hidden");
        }
      }, 120);
    };
    eventNode.addEventListener("mouseenter", show);
    eventNode.addEventListener("mouseleave", hide);
    eventNode.addEventListener("dblclick", () => {
      if (eventNode.dataset.isActive === "1") return;
      editSessionTime(eventNode.dataset.sessionId || "");
    });
  });

  el.timelineList.querySelectorAll(".timeline-event-delete-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteSession(btn.dataset.sessionId || "");
    });
  });
}

function editSessionTime(sessionId) {
  if (!sessionId) return;
  const session = state.sessions.find((s) => s.id === sessionId);
  if (!session) return;
  const startText = prompt("開始日時（YYYY-MM-DD HH:MM）", formatPromptDateTime(session.startAt));
  if (startText === null) return;
  const endText = prompt("終了日時（YYYY-MM-DD HH:MM）", formatPromptDateTime(session.endAt));
  if (endText === null) return;
  const nextStart = parsePromptDateTime(startText.trim());
  const nextEnd = parsePromptDateTime(endText.trim());
  if (!Number.isFinite(nextStart) || !Number.isFinite(nextEnd) || nextEnd <= nextStart) {
    alert("日時形式が不正か、終了が開始以前です。");
    return;
  }
  session.startAt = nextStart;
  session.endAt = nextEnd;
  session.updatedAt = Date.now();
  persistAndRender();
}

function deleteSession(sessionId) {
  if (!sessionId) return;
  const session = state.sessions.find((s) => s.id === sessionId);
  if (!session) return;
  const ok = confirm("このタイムライン記録を削除しますか？");
  if (!ok) return;
  state.sessions = state.sessions.filter((s) => s.id !== sessionId);
  if (!state.deletedSessionIds.includes(sessionId)) {
    state.deletedSessionIds.push(sessionId);
  }
  persistAndRender();
}

function parseTags(input) {
  return input
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

function normalizeStatus(status) {
  return STATUS_OPTIONS.includes(status) ? status : "未着手";
}

function normalizeRecurrence(recurrence) {
  return REPEAT_OPTIONS.includes(recurrence) ? recurrence : "none";
}

function formatRecurrenceLabel(recurrence) {
  const map = {
    none: "なし",
    daily: "毎日",
    weekly: "毎週",
    monthly: "月1",
  };
  return map[normalizeRecurrence(recurrence)];
}

function applyRecurringResets() {
  let changed = false;
  state.tasks.forEach((task) => {
    const recurrence = normalizeRecurrence(task.recurrence);
    task.recurrence = recurrence;
    if (recurrence === "none") {
      task.recurrenceResetKey = "";
      return;
    }
    const currentKey = getRecurrencePeriodKey(recurrence);
    if (!task.recurrenceResetKey) {
      task.recurrenceResetKey = currentKey;
      return;
    }
    if (task.recurrenceResetKey !== currentKey && normalizeStatus(task.status) === "完了") {
      task.status = "未着手";
      task.updatedAt = Date.now();
      changed = true;
    }
    if (task.recurrenceResetKey !== currentKey) {
      task.recurrenceResetKey = currentKey;
      task.updatedAt = Date.now();
      changed = true;
    }
  });
  return changed;
}

function getTodayKey() {
  return formatDateInput(new Date());
}

function getRecurrencePeriodKey(recurrence) {
  const now = new Date();
  if (recurrence === "daily") {
    return formatDateInput(now);
  }
  if (recurrence === "weekly") {
    const monday = new Date(now);
    const diffFromMonday = (monday.getDay() + 6) % 7;
    monday.setDate(monday.getDate() - diffFromMonday);
    return formatDateInput(monday);
  }
  if (recurrence === "monthly") {
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  }
  return "";
}

function normalizeSummaryTab(tab) {
  const allowed = ["task", "mid", "large", "top", "tags"];
  return allowed.includes(tab) ? tab : "task";
}

function formatRangeLabel(range, bounds) {
  const start = new Date(bounds.start);
  const end = new Date(bounds.end);
  end.setDate(end.getDate() - 1);

  if (range === "day") {
    return `${start.getMonth() + 1}/${start.getDate()}`;
  }
  if (range === "week") {
    return `${start.getMonth() + 1}/${start.getDate()}〜${end.getMonth() + 1}/${end.getDate()}`;
  }
  return `${start.getFullYear()}年${start.getMonth() + 1}月 (${start.getMonth() + 1}/${start.getDate()}〜${end.getMonth() + 1}/${end.getDate()})`;
}

function getRangeBounds(range, nowMs, offset = 0) {
  const start = new Date(nowMs);
  const end = new Date(nowMs);

  if (range === "day") {
    start.setHours(0, 0, 0, 0);
    start.setDate(start.getDate() + offset);
    end.setDate(start.getDate() + 1);
    end.setHours(0, 0, 0, 0);
    return { start: start.getTime(), end: end.getTime() };
  }

  if (range === "week") {
    const day = start.getDay();
    const diffFromMonday = (day + 6) % 7;
    start.setDate(start.getDate() - diffFromMonday);
    start.setDate(start.getDate() + offset * 7);
    start.setHours(0, 0, 0, 0);
    end.setTime(start.getTime());
    end.setDate(start.getDate() + 7);
    return { start: start.getTime(), end: end.getTime() };
  }

  start.setDate(1);
  start.setMonth(start.getMonth() + offset);
  start.setHours(0, 0, 0, 0);
  end.setFullYear(start.getFullYear(), start.getMonth() + 1, 1);
  end.setHours(0, 0, 0, 0);
  return { start: start.getTime(), end: end.getTime() };
}

function overlap(aStart, aEnd, bStart, bEnd) {
  return Math.max(0, Math.min(aEnd, bEnd) - Math.max(aStart, bStart));
}

function addDuration(map, key, ms) {
  map.set(key, (map.get(key) || 0) + ms);
}

function formatDuration(ms) {
  const totalMin = Math.floor(ms / 60000);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return `${h}時間 ${m}分`;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./service-worker.js").catch(() => {
      // Ignore SW registration failure to avoid blocking app usage.
    });
  });
}

function persistAndRender(cloudMode = "queued") {
  state.lastDataChangeAt = Date.now();
  persistState();
  renderAll();
  if (cloudMode === "now") {
    queueCloudSave(0);
    return;
  }
  if (cloudMode === "queued") {
    queueCloudSave(1200);
  }
}

function persistUiAndRender() {
  persistState();
  renderAll();
}

function queueCloudSave(delayMs = 1200) {
  cloudSavePending = true;
  if (cloudSaveDebounceTimer) {
    clearTimeout(cloudSaveDebounceTimer);
  }
  cloudSaveDebounceTimer = setTimeout(() => {
    cloudSaveDebounceTimer = null;
    flushQueuedCloudSave();
  }, Math.max(0, delayMs));
}

async function flushQueuedCloudSave() {
  if (!cloudSavePending) return;
  const endpoint = el.cloudEndpoint.value.trim();
  if (!endpoint) {
    cloudSavePending = false;
    return;
  }
  if (isCloudSyncBusy) {
    queueCloudSave(1500);
    return;
  }

  cloudSavePending = false;
  isCloudSyncBusy = true;
  setCloudBusy(true);
  try {
    await cloudSaveRequest(endpoint);
  } catch {
    // Silent failure for background queued sync.
  } finally {
    isCloudSyncBusy = false;
    setCloudBusy(false);
    if (cloudSavePending) {
      queueCloudSave(500);
    }
  }
}

function loadState() {
  const primary = readPersistCandidate(STORAGE_KEY);
  const recovery = readPersistCandidate(STORAGE_RECOVERY_KEY);
  const latest = pickLatestCandidate(primary, recovery);
  if (!latest) return initialState();
  return migrateState(latest.data);
}

function migrateState(parsed) {
  const next = {
    topGroups: Array.isArray(parsed.topGroups)
      ? parsed.topGroups.map((g) => ({
          ...g,
          archived: Boolean(g.archived),
          updatedAt: Number.isFinite(g.updatedAt) ? g.updatedAt : 0,
        }))
      : [],
    largeGroups: Array.isArray(parsed.largeGroups)
      ? parsed.largeGroups.map((g) => ({
          ...g,
          archived: Boolean(g.archived),
          updatedAt: Number.isFinite(g.updatedAt) ? g.updatedAt : 0,
        }))
      : [],
    midGroups: Array.isArray(parsed.midGroups)
      ? parsed.midGroups.map((g) => ({
          ...g,
          archived: Boolean(g.archived),
          updatedAt: Number.isFinite(g.updatedAt) ? g.updatedAt : 0,
        }))
      : [],
    tasks: [],
    sessions: Array.isArray(parsed.sessions)
      ? parsed.sessions.map((s) => ({
          ...s,
          updatedAt: Number.isFinite(s.updatedAt) ? s.updatedAt : 0,
        }))
      : [],
    deletedTaskIds: uniqueStrings(parsed.deletedTaskIds),
    deletedSessionIds: uniqueStrings(parsed.deletedSessionIds),
    activeSession: parsed.activeSession || null,
    taskFilterStatuses: Array.isArray(parsed.taskFilterStatuses)
      ? parsed.taskFilterStatuses.filter((s) => STATUS_OPTIONS.includes(s))
      : [...STATUS_OPTIONS],
    summaryTab: normalizeSummaryTab(parsed.summaryTab),
    summaryOffsets: normalizeSummaryOffsets(parsed.summaryOffsets),
    summaryExpanded: normalizeSummaryExpanded(parsed.summaryExpanded),
    taskParentOrder: Array.isArray(parsed.taskParentOrder) ? parsed.taskParentOrder : [],
    uiSelections: normalizeUiSelections(parsed.uiSelections),
    archiveView: normalizeArchiveView(parsed.archiveView),
    taskLargeFilterValue:
      typeof parsed.taskLargeFilterValue === "string" ? parsed.taskLargeFilterValue : "all",
    taskMidFilterValue:
      typeof parsed.taskMidFilterValue === "string" ? parsed.taskMidFilterValue : "all",
    taskTodayFilterValue: normalizeTaskTodayFilter(parsed.taskTodayFilterValue),
    todayTaskDateKey:
      typeof parsed.todayTaskDateKey === "string" && parsed.todayTaskDateKey
        ? parsed.todayTaskDateKey
        : getTodayKey(),
    manualCollapsed: typeof parsed.manualCollapsed === "boolean" ? parsed.manualCollapsed : true,
    timelineView: normalizeTimelineView(parsed.timelineView),
    addSectionsCollapsed:
      typeof parsed.addSectionsCollapsed === "boolean" ? parsed.addSectionsCollapsed : false,
    lastDataChangeAt:
      Number.isFinite(parsed.lastDataChangeAt) && parsed.lastDataChangeAt > 0
        ? parsed.lastDataChangeAt
        : Date.now(),
  };
  next.tasks = Array.isArray(parsed.tasks)
    ? parsed.tasks.map((task) =>
        normalizeTaskRecord(task, next.midGroups, next.largeGroups)
      )
    : [];
  const deletedTaskSet = new Set(next.deletedTaskIds);
  const deletedSessionSet = new Set(next.deletedSessionIds);
  next.tasks = next.tasks.filter((t) => !deletedTaskSet.has(t.id));
  next.sessions = next.sessions.filter((s) => !deletedSessionSet.has(s.id));

  const hasMissingTopRef = next.largeGroups.some((g) => !g.topGroupId);
  if (!next.topGroups.length || hasMissingTopRef) {
    const fallbackId = "legacy-top";
    if (!next.topGroups.find((g) => g.id === fallbackId)) {
      next.topGroups.unshift({ id: fallbackId, name: "未分類(旧データ)" });
    }
    next.largeGroups = next.largeGroups.map((g) => ({ ...g, topGroupId: g.topGroupId || fallbackId }));
  }

  if (!next.taskFilterStatuses.length) {
    next.taskFilterStatuses = [...STATUS_OPTIONS];
  }

  return next;
}

function persistState() {
  const payload = JSON.stringify({
    savedAt: Date.now(),
    data: state,
  });
  try {
    localStorage.setItem(STORAGE_KEY, payload);
    localStorage.setItem(STORAGE_RECOVERY_KEY, payload);
  } catch {
    // Ignore quota/storage errors to avoid breaking the UI flow.
  }
}

function initialState() {
  return {
    topGroups: [],
    largeGroups: [],
    midGroups: [],
    tasks: [],
    sessions: [],
    deletedTaskIds: [],
    deletedSessionIds: [],
    activeSession: null,
    taskFilterStatuses: [...STATUS_OPTIONS],
    summaryTab: "task",
    summaryOffsets: { day: 0, week: 0, month: 0 },
    summaryExpanded: { task: false, mid: false, large: false, top: false, tags: false },
    taskParentOrder: [],
    uiSelections: { largeTopId: "", midLargeId: "", taskParentValue: "" },
    archiveView: { top: false, large: false, parent: false },
    taskLargeFilterValue: "all",
    taskMidFilterValue: "all",
    taskTodayFilterValue: "all",
    todayTaskDateKey: getTodayKey(),
    manualCollapsed: true,
    timelineView: "calendar",
    addSectionsCollapsed: false,
    lastDataChangeAt: Date.now(),
  };
}

function resolveStateUpdatedAt(target, fallback = 0) {
  const stateStamp =
    target && Number.isFinite(target.lastDataChangeAt) ? target.lastDataChangeAt : 0;
  const fallbackStamp = Number.isFinite(fallback) ? fallback : 0;
  return Math.max(stateStamp, fallbackStamp);
}

async function mergeStateWithCloud(endpoint, localState) {
  try {
    const parsed = await cloudLoadRequest(endpoint);
    if (!parsed.data) return localState;
    const remoteState = migrateState(parsed.data);
    return mergeStates(remoteState, localState);
  } catch {
    return localState;
  }
}

function mergeStates(base, incoming) {
  const deletedTaskIds = uniqueStrings([...(base.deletedTaskIds || []), ...(incoming.deletedTaskIds || [])]);
  const deletedSessionIds = uniqueStrings([
    ...(base.deletedSessionIds || []),
    ...(incoming.deletedSessionIds || []),
  ]);
  const deletedTaskSet = new Set(deletedTaskIds);
  const deletedSessionSet = new Set(deletedSessionIds);

  const baseUpdatedAt = resolveStateUpdatedAt(base);
  const incomingUpdatedAt = resolveStateUpdatedAt(incoming);
  const mergedActiveSession =
    incomingUpdatedAt >= baseUpdatedAt ? incoming.activeSession : base.activeSession;

  const merged = {
    ...base,
    ...incoming,
    topGroups: mergeEntities(base.topGroups, incoming.topGroups),
    largeGroups: mergeEntities(base.largeGroups, incoming.largeGroups),
    midGroups: mergeEntities(base.midGroups, incoming.midGroups),
    tasks: mergeEntities(base.tasks, incoming.tasks).filter((t) => !deletedTaskSet.has(t.id)),
    sessions: mergeEntities(base.sessions, incoming.sessions).filter((s) => !deletedSessionSet.has(s.id)),
    deletedTaskIds,
    deletedSessionIds,
    activeSession: mergedActiveSession || null,
    summaryOffsets: incoming.summaryOffsets || base.summaryOffsets,
    summaryExpanded: incoming.summaryExpanded || base.summaryExpanded,
    uiSelections: incoming.uiSelections || base.uiSelections,
    archiveView: incoming.archiveView || base.archiveView,
    taskFilterStatuses: incoming.taskFilterStatuses || base.taskFilterStatuses,
    taskParentOrder: incoming.taskParentOrder || base.taskParentOrder,
    lastDataChangeAt: Math.max(resolveStateUpdatedAt(base), resolveStateUpdatedAt(incoming)),
  };

  return migrateState(merged);
}

function mergeEntities(baseList, incomingList) {
  const byId = new Map();
  (Array.isArray(baseList) ? baseList : []).forEach((item) => {
    if (item && item.id) byId.set(item.id, item);
  });
  (Array.isArray(incomingList) ? incomingList : []).forEach((item) => {
    if (!item || !item.id) return;
    const current = byId.get(item.id);
    if (!current) {
      byId.set(item.id, item);
      return;
    }
    const currentStamp = Number.isFinite(current.updatedAt) ? current.updatedAt : 0;
    const nextStamp = Number.isFinite(item.updatedAt) ? item.updatedAt : 0;
    byId.set(item.id, nextStamp >= currentStamp ? item : current);
  });
  return [...byId.values()];
}

function uniqueStrings(values) {
  return [...new Set((values || []).filter((v) => typeof v === "string" && v))];
}

function uid() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeSummaryOffsets(offsets) {
  const source = offsets && typeof offsets === "object" ? offsets : {};
  return {
    day: Number.isFinite(source.day) ? Math.min(0, Math.trunc(source.day)) : 0,
    week: Number.isFinite(source.week) ? Math.min(0, Math.trunc(source.week)) : 0,
    month: Number.isFinite(source.month) ? Math.min(0, Math.trunc(source.month)) : 0,
  };
}

function normalizeSummaryExpanded(expanded) {
  const source = expanded && typeof expanded === "object" ? expanded : {};
  return {
    task: Boolean(source.task),
    mid: Boolean(source.mid),
    large: Boolean(source.large),
    top: Boolean(source.top),
    tags: Boolean(source.tags),
  };
}

function normalizeUiSelections(selections) {
  const source = selections && typeof selections === "object" ? selections : {};
  return {
    largeTopId: typeof source.largeTopId === "string" ? source.largeTopId : "",
    midLargeId: typeof source.midLargeId === "string" ? source.midLargeId : "",
    taskParentValue: typeof source.taskParentValue === "string" ? source.taskParentValue : "",
  };
}

function normalizeArchiveView(view) {
  const source = view && typeof view === "object" ? view : {};
  return {
    top: Boolean(source.top),
    large: Boolean(source.large),
    parent: Boolean(source.parent),
  };
}

function normalizeTaskRecord(task, midGroups, largeGroups) {
  const normalized = {
    ...task,
    status: normalizeStatus(task.status),
    isTodayTask: Boolean(task.isTodayTask),
    updatedAt: Number.isFinite(task.updatedAt) ? task.updatedAt : 0,
    recurrence: normalizeRecurrence(task.recurrence),
    recurrenceResetKey:
      typeof task.recurrenceResetKey === "string" ? task.recurrenceResetKey : "",
  };

  if (normalized.parentType === "large") {
    if (!normalized.largeGroupId && normalized.midGroupId) {
      const mid = midGroups.find((m) => m.id === normalized.midGroupId);
      normalized.largeGroupId = mid ? mid.largeGroupId : null;
    }
    normalized.midGroupId = null;
    return normalized;
  }

  if (normalized.parentType === "mid") {
    normalized.largeGroupId = null;
    return normalized;
  }

  if (normalized.midGroupId) {
    normalized.parentType = "mid";
    normalized.largeGroupId = null;
    return normalized;
  }

  if (normalized.largeGroupId) {
    normalized.parentType = "large";
    normalized.midGroupId = null;
    return normalized;
  }

  if (largeGroups.length > 0) {
    normalized.parentType = "large";
    normalized.largeGroupId = largeGroups[0].id;
    normalized.midGroupId = null;
    return normalized;
  }

  normalized.parentType = "mid";
  normalized.midGroupId = midGroups[0] ? midGroups[0].id : null;
  normalized.largeGroupId = null;
  return normalized;
}

function parseTaskParentValue(value) {
  if (!value) return null;
  if (value.startsWith("mid:")) {
    return { parentType: "mid", midGroupId: value.slice(4) };
  }
  if (value.startsWith("large:")) {
    return { parentType: "large", largeGroupId: value.slice(6) };
  }
  return { parentType: "mid", midGroupId: value };
}

function normalizeTaskTodayFilter(value) {
  return value === "today" ? "today" : "all";
}

function normalizeTimelineView(value) {
  return value === "list" ? "list" : "calendar";
}

function getTaskParentOptions() {
  const options = [
    ...state.largeGroups.map((g) => {
      const top = state.topGroups.find((t) => t.id === g.topGroupId);
      const topName = top ? top.name : "(不明な最上位グループ)";
      return {
        value: `large:${g.id}`,
        label: `[中なし] ${topName} > ${g.name}`,
        archived: Boolean(g.archived),
        entityType: "large",
        entityId: g.id,
      };
    }),
    ...state.midGroups.map((m) => {
      const large = state.largeGroups.find((g) => g.id === m.largeGroupId);
      const top = large ? state.topGroups.find((t) => t.id === large.topGroupId) : null;
      const topName = top ? top.name : "(不明な最上位グループ)";
      const largeName = large ? large.name : "(不明な大グループ)";
      return {
        value: `mid:${m.id}`,
        label: `${topName} > ${largeName} > ${m.name}`,
        archived: Boolean(m.archived),
        entityType: "mid",
        entityId: m.id,
      };
    }),
  ];

  const order = Array.isArray(state.taskParentOrder) ? state.taskParentOrder : [];
  const byValue = new Map(options.map((item) => [item.value, item]));
  const merged = [
    ...order.filter((value) => byValue.has(value)),
    ...options.map((item) => item.value).filter((value) => !order.includes(value)),
  ];
  return merged.map((value) => byValue.get(value));
}

function ensureSelection(currentValue, options) {
  if (options.length === 0) return "";
  if (options.some((item) => item.value === currentValue)) return currentValue;
  return options[0].value;
}

function reorderByValues(array, orderedValues) {
  const byId = new Map(array.map((item) => [item.id, item]));
  const orderedSet = new Set(orderedValues);
  const reorderedHead = orderedValues.map((id) => byId.get(id)).filter(Boolean);
  const tail = array.filter((item) => !orderedSet.has(item.id));
  array.splice(0, array.length, ...reorderedHead, ...tail);
}

function initDragDropdownGlobalClose() {
  if (dragDropdownListenerInitialized) return;
  document.addEventListener("click", (e) => {
    document.querySelectorAll(".drag-dropdown.open").forEach((node) => {
      if (!node.contains(e.target)) node.classList.remove("open");
    });
  });
  dragDropdownListenerInitialized = true;
}

function renderDragDropdown(
  container,
  options,
  selectedValue,
  emptyLabel,
  showArchived,
  onToggleShowArchived,
  onSelect,
  onReorder,
  onToggleArchive
) {
  const visible = filterArchivedOptions(options, showArchived);
  if (!visible.length) {
    const help = options.some((item) => item.archived)
      ? `<button type="button" class="dd-archive-toggle">${showArchived ? "アーカイブを隠す" : "アーカイブを表示"}</button>`
      : "";
    container.innerHTML = `<button type="button" class="dd-trigger" disabled>${escapeHtml(
      emptyLabel
    )}</button>${help}`;
    container.classList.remove("open");
    const toggleBtn = container.querySelector(".dd-archive-toggle");
    if (toggleBtn) {
      toggleBtn.addEventListener("click", onToggleShowArchived);
    }
    return;
  }

  const selected = visible.find((item) => item.value === selectedValue) || visible[0];
  container.innerHTML = `
    <button type="button" class="dd-trigger">
      <span class="dd-value">${escapeHtml(selected.label)}</span>
      <span class="dd-arrow">▾</span>
    </button>
    <div class="dd-menu">
      <div class="dd-menu-head">
        <button type="button" class="dd-archive-toggle">${showArchived ? "アーカイブを隠す" : "アーカイブを表示"}</button>
      </div>
      ${visible
        .map(
          (item) => `<div class="dd-item" draggable="true" data-value="${escapeHtml(item.value)}">
        <span class="dd-handle">≡</span>
        <span class="dd-label">${escapeHtml(item.label)}</span>
        <button type="button" class="dd-archive-btn" data-entity-type="${item.entityType || ""}" data-entity-id="${item.entityId || ""}">
          ${item.archived ? "復帰" : "アーカイブ"}
        </button>
      </div>`
        )
        .join("")}
    </div>
  `;

  const trigger = container.querySelector(".dd-trigger");
  const items = [...container.querySelectorAll(".dd-item")];
  const archiveButtons = [...container.querySelectorAll(".dd-archive-btn")];
  const archiveToggleBtn = container.querySelector(".dd-archive-toggle");
  let dragValue = "";

  trigger.addEventListener("click", () => {
    container.classList.toggle("open");
  });

  items.forEach((item) => {
    item.addEventListener("click", () => {
      onSelect(item.dataset.value);
      container.classList.remove("open");
    });
    item.addEventListener("dragstart", () => {
      dragValue = item.dataset.value;
    });
    item.addEventListener("dragover", (e) => {
      e.preventDefault();
      item.classList.add("drag-over");
    });
    item.addEventListener("dragleave", () => {
      item.classList.remove("drag-over");
    });
    item.addEventListener("drop", (e) => {
      e.preventDefault();
      item.classList.remove("drag-over");
      const targetValue = item.dataset.value;
      if (!dragValue || dragValue === targetValue) return;
      const ordered = moveValueInArray(
        visible.map((opt) => opt.value),
        dragValue,
        targetValue
      );
      onReorder(ordered);
    });
  });

  archiveButtons.forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const type = btn.dataset.entityType;
      const id = btn.dataset.entityId;
      if (type && id) onToggleArchive(type, id);
    });
  });

  if (archiveToggleBtn) {
    archiveToggleBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      onToggleShowArchived();
    });
  }
}

function moveValueInArray(values, fromValue, toValue) {
  const next = [...values];
  const fromIndex = next.indexOf(fromValue);
  const toIndex = next.indexOf(toValue);
  if (fromIndex < 0 || toIndex < 0) return next;
  const [item] = next.splice(fromIndex, 1);
  next.splice(toIndex, 0, item);
  return next;
}

function toggleArchive(entityType, entityId) {
  const map = {
    top: state.topGroups,
    large: state.largeGroups,
    mid: state.midGroups,
  };
  const targetArray = map[entityType];
  if (!targetArray) return;
  const target = targetArray.find((item) => item.id === entityId);
  if (!target) return;
  target.archived = !target.archived;
  target.updatedAt = Date.now();
}

function filterArchivedOptions(options, showArchived) {
  if (showArchived) return options.filter((item) => item.archived);
  return options.filter((item) => !item.archived);
}

function resolveVisibleOptions(options, key) {
  const showArchived = Boolean(state.archiveView[key]);
  let visible = filterArchivedOptions(options, showArchived);
  if (showArchived && visible.length === 0) {
    state.archiveView[key] = false;
    visible = filterArchivedOptions(options, false);
  }
  return visible;
}

function startAutoSave() {
  setInterval(() => {
    persistState();
  }, AUTO_SAVE_INTERVAL_MS);
}

function readPersistCandidate(key) {
  const raw = localStorage.getItem(key);
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object" && parsed.data) {
      return {
        savedAt: Number.isFinite(parsed.savedAt) ? parsed.savedAt : 0,
        data: parsed.data,
      };
    }
    return { savedAt: 0, data: parsed };
  } catch {
    return null;
  }
}

function pickLatestCandidate(a, b) {
  if (!a && !b) return null;
  if (!a) return b;
  if (!b) return a;
  return a.savedAt >= b.savedAt ? a : b;
}

function isStateEmpty(target) {
  return (
    target.topGroups.length === 0 &&
    target.largeGroups.length === 0 &&
    target.midGroups.length === 0 &&
    target.tasks.length === 0 &&
    target.sessions.length === 0 &&
    !target.activeSession
  );
}

function formatDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatPromptDateTime(ms) {
  const d = new Date(ms);
  const date = formatDateInput(d);
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${date} ${hh}:${mm}`;
}

function parsePromptDateTime(text) {
  const normalized = text.replace(/\//g, "-").replace("T", " ");
  const match = normalized.match(
    /^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})$/
  );
  if (!match) return NaN;
  const [, y, m, d, hh, mm] = match;
  const date = new Date(
    Number(y),
    Number(m) - 1,
    Number(d),
    Number(hh),
    Number(mm),
    0,
    0
  );
  const valid =
    date.getFullYear() === Number(y) &&
    date.getMonth() === Number(m) - 1 &&
    date.getDate() === Number(d) &&
    date.getHours() === Number(hh) &&
    date.getMinutes() === Number(mm);
  return valid ? date.getTime() : NaN;
}

function formatTimelineDay(ms) {
  return new Date(ms).toLocaleDateString("ja-JP", {
    year: "numeric",
    month: "numeric",
    day: "numeric",
    weekday: "short",
  });
}

function formatTimelineTime(ms) {
  return new Date(ms).toLocaleTimeString("ja-JP", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
}

function formatTimelineDateTime(ms) {
  const d = new Date(ms);
  const dateLabel = `${d.getMonth() + 1}/${d.getDate()}`;
  return `${dateLabel} ${formatTimelineTime(ms)}`;
}
