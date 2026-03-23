// ==UserScript==
// @name         OWA Calendar Notifier
// @namespace    https://github.com/owa-notifications
// @version      1.0
// @description  Уведомления о предстоящих встречах в OWA (Exchange 2019 on-premises)
// @match        https://*/owa/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ─── настройки ──────────────────────────────────────────────────────────
  const POLL_INTERVAL = 30_000;   // 30 секунд
  const NOTIFY_BEFORE_MIN = 15;   // уведомлять за 15 мин до начала
  const CALENDAR_WINDOW_MIN = 60; // окно просмотра — 60 мин вперёд

  // ─── состояние ──────────────────────────────────────────────────────────
  let isRunning = false;
  let pollTimer = null;
  const notifiedEventIds = new Set();
  const imminentTimers = new Map(); // eventId → timerId для 5-секундных напоминаний
  const threeMinTimers = new Map(); // eventId → timerId для 3-минутных напоминаний
  let notifCount = 0;
  let eventCount = 0;
  let isCollapsed = true;
  let firstErrorTime = null;
  const DISCONNECT_TIMEOUT = 5 * 60 * 1000; // 5 минут
  let disconnectNotified = false;

  // ─── UI: floating-виджет ────────────────────────────────────────────────
  function createWidget() {
    const widget = document.createElement('div');
    widget.id = 'owa-notifier-widget';
    widget.innerHTML = `
      <style>
        #owa-notifier-widget {
          position: fixed;
          bottom: 16px;
          right: 16px;
          z-index: 999999;
          font-family: 'Segoe UI', system-ui, sans-serif;
          font-size: 13px;
          background: #1e2130;
          color: #e2e8f0;
          border: 1px solid #2d3148;
          border-radius: 10px;
          box-shadow: 0 4px 24px rgba(0,0,0,0.4);
          min-width: 260px;
          max-width: 320px;
          overflow: hidden;
          transition: opacity 0.2s;
          user-select: none;
        }
        #owa-notifier-widget .onw-log {
          user-select: text;
        }
        #owa-notifier-widget * { box-sizing: border-box; margin: 0; padding: 0; }
        #owa-notifier-header {
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 10px 14px;
          cursor: pointer;
          background: #161825;
          border-bottom: 1px solid #2d3148;
        }
        #owa-notifier-header:hover { background: #1a1d2e; }
        #owa-notifier-dot {
          width: 8px; height: 8px;
          border-radius: 50%;
          background: #4a4a5a;
          flex-shrink: 0;
          transition: background 0.3s, box-shadow 0.3s;
        }
        #owa-notifier-dot.active { background: #22d3a5; box-shadow: 0 0 6px #22d3a588; }
        #owa-notifier-dot.error  { background: #f87171; box-shadow: 0 0 6px #f8717188; }
        #owa-notifier-title {
          font-size: 11px;
          font-weight: 600;
          color: #94a3b8;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          flex: 1;
        }
        #owa-notifier-badge {
          font-size: 11px;
          font-weight: 600;
          color: #0f1117;
          background: #22d3a5;
          border-radius: 10px;
          padding: 1px 7px;
          min-width: 20px;
          text-align: center;
        }
        #owa-notifier-toggle {
          font-size: 14px;
          color: #64748b;
          transition: transform 0.2s;
        }
        #owa-notifier-toggle.collapsed { transform: rotate(180deg); }
        #owa-notifier-body {
          padding: 10px 14px;
          display: flex;
          flex-direction: column;
          gap: 8px;
        }
        #owa-notifier-body.hidden { display: none; }
        .onw-status {
          font-size: 12px;
          color: #94a3b8;
          line-height: 1.4;
        }
        .onw-event {
          background: #13151f;
          border-radius: 6px;
          padding: 8px 10px;
          display: flex;
          flex-direction: column;
          gap: 3px;
        }
        .onw-event-subject {
          font-size: 13px;
          font-weight: 600;
          color: #e2e8f0;
        }
        .onw-event-time {
          font-size: 11px;
          color: #94a3b8;
        }
        .onw-event-location {
          font-size: 11px;
          color: #64748b;
        }
        .onw-event-countdown {
          font-size: 12px;
          font-weight: 600;
          color: #fbbf24;
          margin-top: 2px;
        }
        .onw-controls {
          display: flex;
          gap: 6px;
        }
        .onw-btn {
          flex: 1;
          padding: 6px 8px;
          border-radius: 6px;
          border: 1px solid #2d3148;
          background: #161825;
          color: #94a3b8;
          font-size: 11px;
          font-weight: 600;
          cursor: pointer;
          text-align: center;
        }
        .onw-btn:hover { background: #1a1d2e; }
        .onw-btn.primary {
          background: #22d3a5;
          color: #0f1117;
          border-color: #22d3a5;
        }
        .onw-btn.primary:hover { opacity: 0.9; }
        .onw-log {
          max-height: 120px;
          overflow-y: auto;
          display: flex;
          flex-direction: column;
          gap: 2px;
          scrollbar-width: thin;
          scrollbar-color: #2d3148 transparent;
        }
        .onw-log::-webkit-scrollbar { width: 4px; }
        .onw-log::-webkit-scrollbar-track { background: transparent; }
        .onw-log::-webkit-scrollbar-thumb { background: #2d3148; border-radius: 2px; }
        .onw-log::-webkit-scrollbar-thumb:hover { background: #3d4168; }
        .onw-log-entry {
          font-size: 10px;
          font-family: 'Cascadia Code', 'Consolas', monospace;
          color: #475569;
          line-height: 1.3;
        }
        .onw-log-entry.new  { color: #22d3a5; }
        .onw-log-entry.err  { color: #f87171; }
        .onw-log-entry.warn { color: #fbbf24; }
      </style>
      <div id="owa-notifier-header">
        <div id="owa-notifier-dot"></div>
        <span id="owa-notifier-title">OWA Notifier</span>
        <span id="owa-notifier-badge">0</span>
        <span id="owa-notifier-toggle" class="collapsed">▾</span>
      </div>
      <div id="owa-notifier-body" class="hidden">
        <div class="onw-status" id="onw-status">Инициализация...</div>
        <div id="onw-next-event"></div>
        <div class="onw-controls">
          <button class="onw-btn primary" id="onw-btn-toggle">Запустить</button>
          <button class="onw-btn" id="onw-btn-notif">Разрешить уведомления</button>
          <button class="onw-btn" id="onw-btn-test">Тест</button>
        </div>
        <div class="onw-log" id="onw-log"></div>
      </div>
    `;
    document.body.appendChild(widget);

    // свернуть/развернуть
    document.getElementById('owa-notifier-header').addEventListener('click', () => {
      isCollapsed = !isCollapsed;
      document.getElementById('owa-notifier-body').classList.toggle('hidden', isCollapsed);
      document.getElementById('owa-notifier-toggle').classList.toggle('collapsed', isCollapsed);
    });

    document.getElementById('onw-btn-toggle').addEventListener('click', togglePolling);
    document.getElementById('onw-btn-notif').addEventListener('click', requestNotificationPermission);
    document.getElementById('onw-btn-test').addEventListener('click', () => {
      addLog('Тестовое уведомление', 'new');
      showNotification(
        'Тестовая встреча',
        'Переговорная 3.14\nНажмите чтобы присоединиться\nЧерез 5 мин',
        null
      );
    });
  }

  // ─── UI хелперы ─────────────────────────────────────────────────────────
  function setDot(state) {
    const dot = document.getElementById('owa-notifier-dot');
    if (dot) dot.className = state || '';
  }

  function setStatus(text, dotState) {
    const el = document.getElementById('onw-status');
    if (el) el.textContent = text;
    setDot(dotState);
  }

  function addLog(text, type) {
    const log = document.getElementById('onw-log');
    if (!log) return;
    const entry = document.createElement('div');
    entry.className = 'onw-log-entry' + (type ? ' ' + type : '');
    const time = new Date().toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    entry.textContent = `[${time}] ${text}`;
    log.prepend(entry);
    while (log.children.length > 30) log.lastChild.remove();
  }

  function updateBadge() {
    const badge = document.getElementById('owa-notifier-badge');
    if (badge) badge.textContent = eventCount;
  }

  function showNextEvent(event) {
    const container = document.getElementById('onw-next-event');
    if (!container) return;

    if (!event) {
      container.innerHTML = '';
      return;
    }

    const startTime = new Date(event.start);
    const endTime = new Date(event.end);
    const timeFmt = { hour: '2-digit', minute: '2-digit' };
    const timeStr = startTime.toLocaleTimeString('ru-RU', timeFmt) + ' — ' + endTime.toLocaleTimeString('ru-RU', timeFmt);
    const minutesLeft = Math.round((startTime - new Date()) / 60000);
    const countdown = minutesLeft <= 0 ? 'Уже началось' : `Через ${minutesLeft} мин`;

    container.innerHTML = `
      <div class="onw-event">
        <span class="onw-event-subject">${escapeHtml(event.subject)}</span>
        <span class="onw-event-time">${escapeHtml(timeStr)}</span>
        ${event.location ? `<span class="onw-event-location">${escapeHtml(event.location)}</span>` : ''}
        <span class="onw-event-countdown">${escapeHtml(countdown)}</span>
      </div>
    `;
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // ─── уведомления ────────────────────────────────────────────────────────
  async function requestNotificationPermission() {
    if (!('Notification' in window)) {
      addLog('Notification API не поддерживается', 'err');
      return;
    }
    const perm = await Notification.requestPermission();
    if (perm === 'granted') {
      addLog('Разрешение на уведомления получено', 'new');
      updateNotifButton();
    } else {
      addLog('Разрешение на уведомления отклонено', 'err');
    }
  }

  function updateNotifButton() {
    const btn = document.getElementById('onw-btn-notif');
    if (!btn) return;
    if (Notification.permission === 'granted') {
      btn.textContent = `Уведомлений: ${notifCount}`;
      btn.disabled = true;
      btn.style.opacity = '0.5';
    } else if (Notification.permission === 'denied') {
      btn.textContent = 'Заблокировано';
      btn.disabled = true;
      btn.style.opacity = '0.5';
    }
  }

  // ─── извлечение ссылки на встречу ─────────────────────────────────────
  const MEETING_URL_PATTERNS = [
    /https?:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s<>"')]+/i,
    /https?:\/\/meet\.google\.com\/[a-z\-]+/i,
    /https?:\/\/[a-z0-9.-]*zoom\.us\/j\/[^\s<>"')]+/i,
    /https?:\/\/[a-z0-9.-]*ktalk\.ru\/[^\s<>"')]+/i,
  ];

  function extractMeetingUrl(text) {
    if (!text) return null;
    for (const pattern of MEETING_URL_PATTERNS) {
      const match = text.match(pattern);
      if (match) return match[0];
    }
    // fallback: любая https ссылка из body
    const anyUrl = text.match(/https?:\/\/[^\s<>"')]+/i);
    return anyUrl ? anyUrl[0] : null;
  }

  function playNotificationSound() {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.frequency.value = 880;
    gain.gain.value = 0.3;
    osc.start();
    gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.3);
    osc.stop(ctx.currentTime + 0.3);
    // второй тон
    const osc2 = ctx.createOscillator();
    const gain2 = ctx.createGain();
    osc2.connect(gain2);
    gain2.connect(ctx.destination);
    osc2.frequency.value = 1100;
    gain2.gain.value = 0.3;
    osc2.start(ctx.currentTime + 0.15);
    gain2.gain.setValueAtTime(0.3, ctx.currentTime + 0.15);
    gain2.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.5);
    osc2.stop(ctx.currentTime + 0.5);
    setTimeout(() => ctx.close(), 600);
  }

  function showNotification(title, body, meetingUrl) {
    if (Notification.permission !== 'granted') return;
    playNotificationSound();
    const n = new Notification(title, {
      body,
      icon: 'https://res.cdn.office.net/files/fabric-cdn-prod_20230815.002/assets/brand-icons/product/svg/outlook_32x1.svg',
      tag: 'owa-notifier-' + Date.now(),
    });
    n.onclick = () => {
      n.close();
      if (meetingUrl) {
        window.open(meetingUrl, '_blank');
      } else {
        window.focus();
      }
    };
    notifCount++;
    updateNotifButton();
  }

  // ─── определение таймзоны (IANA → Windows) ─────────────────────────────
  function getWindowsTimeZone() {
    const map = {
      'Europe/Moscow': 'Russian Standard Time',
      'Europe/Kaliningrad': 'Kaliningrad Standard Time',
      'Europe/Samara': 'Russia Time Zone 3',
      'Asia/Yekaterinburg': 'Ekaterinburg Standard Time',
      'Asia/Omsk': 'Omsk Standard Time',
      'Asia/Novosibirsk': 'N. Central Asia Standard Time',
      'Asia/Barnaul': 'Altai Standard Time',
      'Asia/Krasnoyarsk': 'North Asia Standard Time',
      'Asia/Irkutsk': 'North Asia East Standard Time',
      'Asia/Yakutsk': 'Yakutsk Standard Time',
      'Asia/Vladivostok': 'Vladivostok Standard Time',
      'Asia/Magadan': 'Magadan Standard Time',
      'Asia/Kamchatka': 'Russia Time Zone 11',
      'Europe/London': 'GMT Standard Time',
      'Europe/Berlin': 'W. Europe Standard Time',
      'America/New_York': 'Eastern Standard Time',
      'America/Chicago': 'Central Standard Time',
      'America/Los_Angeles': 'Pacific Standard Time',
      'Asia/Tokyo': 'Tokyo Standard Time',
      'Asia/Shanghai': 'China Standard Time',
    };
    const iana = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return map[iana] || 'Russian Standard Time';
  }

  // ─── получение canary-токена из cookies ──────────────────────────────────
  function getCanary() {
    const match = document.cookie.match(/X-OWA-CANARY=([^;]+)/);
    return match ? match[1] : null;
  }

  // ─── запрос календаря через GetReminders (раскрывает периодические) ─────
  function formatLocalDateTime(date) {
    return date.getFullYear() + '-' +
      String(date.getMonth() + 1).padStart(2, '0') + '-' +
      String(date.getDate()).padStart(2, '0') + 'T' +
      String(date.getHours()).padStart(2, '0') + ':' +
      String(date.getMinutes()).padStart(2, '0') + ':' +
      String(date.getSeconds()).padStart(2, '0');
  }

  function buildGetRemindersUrl() {
    const now = new Date();
    const later = new Date(now.getTime() + CALENDAR_WINDOW_MIN * 60 * 1000);

    const payload = {
      __type: 'GetRemindersJsonRequest:#Exchange',
      Header: {
        __type: 'JsonRequestHeaders:#Exchange',
        RequestServerVersion: 'Exchange2013',
        TimeZoneContext: {
          __type: 'TimeZoneContext:#Exchange',
          TimeZoneDefinition: {
            __type: 'TimeZoneDefinitionType:#Exchange',
            Id: getWindowsTimeZone(),
          },
        },
      },
      Body: {
        __type: 'GetRemindersRequest:#Exchange',
        EndTime: formatLocalDateTime(later),
        MaxItems: 0,
      },
    };

    return encodeURIComponent(JSON.stringify(payload));
  }

  function parseRemindersResponse(data) {
    const body = data?.Body;
    if (!body) throw new Error('Пустой ответ от сервера');

    if (body.ResponseClass === 'Error') {
      throw new Error(body.MessageText || 'Ошибка GetReminders');
    }

    const reminders = body.Reminders || [];
    const result = [];
    const now = new Date();
    const windowEnd = new Date(now.getTime() + CALENDAR_WINDOW_MIN * 60 * 1000);

    for (const r of reminders) {
      const id = r.ItemId?.Id;
      const subject = r.Subject || '(без темы)';
      const location = r.Location || '';
      if (!id || !r.StartDate) continue;

      let startDate = new Date(r.StartDate);
      let endDate = new Date(r.EndDate || r.StartDate);
      const reminderDate = new Date(r.ReminderTime || r.StartDate);

      // Баг Exchange: для некоторых периодических встреч StartDate/EndDate
      // указывают на прошлое вхождение, а ReminderTime — на актуальное.
      // Если расхождение > 24ч, пересчитываем даты от ReminderTime.
      if (reminderDate - startDate > 24 * 60 * 60 * 1000) {
        const duration = endDate - startDate;
        // По умолчанию напоминание за 15 мин до начала
        startDate = new Date(reminderDate.getTime() + 15 * 60 * 1000);
        endDate = new Date(startDate.getTime() + duration);
      }

      if (endDate < now || startDate > windowEnd) continue;

      result.push({ id, subject, start: startDate.toISOString(), end: endDate.toISOString(), location, meetingUrl: null });
    }

    result.sort((a, b) => new Date(a.start) - new Date(b.start));
    return result;
  }

  async function fetchCalendar() {
    const canary = getCanary();
    if (!canary) throw new Error('X-OWA-CANARY не найден — перезагрузите OWA');

    const response = await fetch('/owa/service.svc?action=GetReminders&EP=1&ID=-11&AC=1', {
      method: 'POST',
      credentials: 'same-origin',
      headers: {
        'Content-Type': 'application/json; charset=UTF-8',
        'Content-Length': '0',
        'Action': 'GetReminders',
        'X-OWA-CANARY': canary,
        'X-OWA-ActionName': 'GetRemindersAction',
        'X-OWA-UrlPostData': buildGetRemindersUrl(),
        'X-Requested-With': 'XMLHttpRequest',
      },
      body: '',
    });

    if (response.status === 401 || response.status === 440) {
      throw new Error('Сессия истекла — перезагрузите OWA');
    }
    if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);

    const data = await response.json();
    return parseRemindersResponse(data);
  }

  // ─── GetItem: получить Body события для извлечения ссылки ──────────────
  function buildGetItemBody(itemId) {
    return JSON.stringify({
      __type: 'GetItemJsonRequest:#Exchange',
      Header: {
        __type: 'JsonRequestHeaders:#Exchange',
        RequestServerVersion: 'Exchange2013',
      },
      Body: {
        __type: 'GetItemRequest:#Exchange',
        ItemShape: {
          __type: 'ItemResponseShape:#Exchange',
          BaseShape: 'IdOnly',
          BodyType: 'Text',
          AdditionalProperties: [
            { __type: 'PropertyUri:#Exchange', FieldURI: 'Body' },
            { __type: 'PropertyUri:#Exchange', FieldURI: 'Location' },
          ],
        },
        ItemIds: [
          { __type: 'ItemId:#Exchange', Id: itemId },
        ],
      },
    });
  }

  async function fetchItemBody(itemId) {
    const canary = getCanary();
    if (!canary) return null;

    const response = await fetch('/owa/service.svc?action=GetItem', {
      method: 'POST',
      credentials: 'same-origin',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Action': 'GetItem',
        'X-OWA-CANARY': canary,
        'X-OWA-ActionName': 'OWANotifierGetItem',
      },
      body: buildGetItemBody(itemId),
    });

    if (!response.ok) return null;

    const data = await response.json();
    const items = data?.Body?.ResponseMessages?.Items;
    if (!items || items.length === 0) return null;

    const msg = items[0];
    if (msg.ResponseClass === 'Error') return null;

    const item = msg.Items?.[0];
    if (!item) return null;

    return {
      body: item.Body?.Value || '',
      location: item.Location || '',
    };
  }

  // кэш ссылок, чтобы не делать GetItem каждый поллинг
  const meetingUrlCache = new Map(); // eventId → meetingUrl | ''

  async function enrichEventWithMeetingUrl(event) {
    if (meetingUrlCache.has(event.id)) {
      event.meetingUrl = meetingUrlCache.get(event.id) || null;
      return;
    }

    // сначала проверяем Location из FindItem
    const urlFromLocation = extractMeetingUrl(event.location);
    if (urlFromLocation) {
      event.meetingUrl = urlFromLocation;
      meetingUrlCache.set(event.id, urlFromLocation);
      return;
    }

    // делаем GetItem для получения Body
    const details = await fetchItemBody(event.id);
    if (details) {
      const url = extractMeetingUrl(details.location) || extractMeetingUrl(details.body);
      event.meetingUrl = url;
      meetingUrlCache.set(event.id, url || '');
    } else {
      meetingUrlCache.set(event.id, '');
    }
  }

  // ─── очистка notifiedEventIds от прошедших встреч ───────────────────────
  function cleanupNotifiedIds(events) {
    const currentIds = new Set(events.map(e => e.id));
    for (const id of notifiedEventIds) {
      if (!currentIds.has(id)) {
        notifiedEventIds.delete(id);
        meetingUrlCache.delete(id);
      }
    }
  }

  // ─── polling ────────────────────────────────────────────────────────────
  async function poll() {
    if (!isRunning) return;
    lastPollTime = Date.now();
    setStatus('Проверяю календарь...', 'active');

    try {
      const events = await fetchCalendar();
      eventCount = events.length;
      updateBadge();
      cleanupNotifiedIds(events);
      showNextEvent(events.length > 0 ? events[0] : null);

      const now = new Date();
      let notifiedThisRound = 0;

      for (const event of events) {
        const startTime = new Date(event.start);
        const secondsLeft = (startTime - now) / 1000;

        // обогащаем ссылкой на встречу только события в пределах ~16 мин
        if (secondsLeft <= NOTIFY_BEFORE_MIN * 60 + 59 && secondsLeft > 0) {
          await enrichEventWithMeetingUrl(event);
        }

        const timeStr = startTime.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });

        if (secondsLeft <= NOTIFY_BEFORE_MIN * 60 + 59 && secondsLeft > 0 && !notifiedEventIds.has(event.id)) {
          notifiedEventIds.add(event.id);
          addLog(`Встреча в ${timeStr}: ${event.subject}`, 'new');
          showNotification(
            event.subject,
            `${event.location ? event.location + (event.meetingUrl ? '\nНажмите чтобы присоединиться' : '') + '\n' : (event.meetingUrl ? event.meetingUrl.length > 40 ? event.meetingUrl.slice(0, 37) + '...' : event.meetingUrl + '\n' : '')}Начало в ${timeStr}`,
            event.meetingUrl
          );
          notifiedThisRound++;
        }

        // запланировать напоминание за 5 минут до начала
        const msUntil5min = (startTime - now) - 5 * 60 * 1000;
        if (msUntil5min > 0 && msUntil5min < POLL_INTERVAL * 2 && !threeMinTimers.has(event.id)) {
          const ev = event;
          const timerId3 = setTimeout(() => {
            threeMinTimers.delete(ev.id);
            addLog(`Встреча через 5 мин: ${ev.subject}`, 'new');
            showNotification(
              ev.subject,
              `${ev.location ? ev.location + (ev.meetingUrl ? '\nНажмите чтобы присоединиться' : '') + '\n' : (ev.meetingUrl ? ev.meetingUrl.length > 40 ? ev.meetingUrl.slice(0, 37) + '...' : ev.meetingUrl + '\n' : '')}Начало в ${timeStr}`,
              ev.meetingUrl
            );
          }, msUntil5min);
          threeMinTimers.set(event.id, timerId3);
        }

        // запланировать напоминание за 1 минуту до начала
        const msUntilImminent = (startTime - now) - 60 * 1000;
        if (msUntilImminent > 0 && msUntilImminent < POLL_INTERVAL * 2 && !imminentTimers.has(event.id)) {
          const ev = event;
          const timerId = setTimeout(() => {
            imminentTimers.delete(ev.id);
            addLog(`Встреча через 1 мин: ${ev.subject}`, 'new');
            showNotification(
              ev.subject,
              `${ev.location ? ev.location + (ev.meetingUrl ? '\nНажмите чтобы присоединиться' : '') + '\n' : (ev.meetingUrl ? ev.meetingUrl.length > 40 ? ev.meetingUrl.slice(0, 37) + '...' : ev.meetingUrl + '\n' : '')}Начинается!`,
              ev.meetingUrl
            );
          }, msUntilImminent);
          imminentTimers.set(event.id, timerId);
        }
      }

      if (notifiedThisRound === 0) {
        addLog(`Встреч в ближайший час: ${events.length}`);
      }

      setStatus(`Активен · ${events.length} встреч · ${new Date().toLocaleTimeString('ru-RU')}`, 'active');
      firstErrorTime = null;
      disconnectNotified = false;
    } catch (e) {
      if (!firstErrorTime) firstErrorTime = new Date();
      const errorDuration = Math.round((new Date() - firstErrorTime) / 60000);
      addLog(`Ошибка (${errorDuration} мин): ${e.message}`, 'err');
      setStatus(`Нет связи · ${errorDuration} мин`, 'error');

      if ((new Date() - firstErrorTime) >= DISCONNECT_TIMEOUT && !disconnectNotified) {
        disconnectNotified = true;
        showNotification(
          'OWA Notifier: нет связи',
          `Нет подключения уже ${errorDuration} мин.\nУведомления о встречах не приходят!`,
          null
        );
        addLog('Отправлено уведомление о потере связи', 'warn');
      }
    }

    scheduleNext();
  }

  function scheduleNext() {
    if (isRunning) {
      pollTimer = setTimeout(poll, POLL_INTERVAL);
    }
  }

  // ─── пробуждение: проверка при выходе из сна/блокировки ─────────────────
  let lastPollTime = Date.now();

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && isRunning) {
      const elapsed = Date.now() - lastPollTime;
      if (elapsed > POLL_INTERVAL * 1.5) {
        addLog(`Пробуждение после ${Math.round(elapsed / 60000)} мин — проверяю`, 'warn');
        clearTimeout(pollTimer);
        poll();
      }
    }
  });

  // fallback: setInterval ловит случаи когда visibilitychange не сработал
  setInterval(() => {
    if (!isRunning) return;
    const elapsed = Date.now() - lastPollTime;
    if (elapsed > POLL_INTERVAL * 2) {
      addLog(`Пропущен поллинг (${Math.round(elapsed / 60000)} мин) — проверяю`, 'warn');
      clearTimeout(pollTimer);
      poll();
    }
  }, 5000);

  function startPolling() {
    if (isRunning) return;
    isRunning = true;
    notifiedEventIds.clear();
    addLog('Мониторинг запущен', 'new');
    setStatus('Запускаю...', 'active');
    const btn = document.getElementById('onw-btn-toggle');
    if (btn) { btn.textContent = 'Остановить'; btn.classList.add('primary'); }
    poll();
  }

  function stopPolling() {
    isRunning = false;
    clearTimeout(pollTimer);
    for (const timerId of imminentTimers.values()) clearTimeout(timerId);
    imminentTimers.clear();
    for (const timerId of threeMinTimers.values()) clearTimeout(timerId);
    threeMinTimers.clear();
    addLog('Мониторинг остановлен', 'warn');
    setStatus('Остановлен', '');
    setDot('');
    const btn = document.getElementById('onw-btn-toggle');
    if (btn) { btn.textContent = 'Запустить'; btn.classList.remove('primary'); }
  }

  function togglePolling() {
    if (isRunning) stopPolling();
    else startPolling();
  }

  // ─── инициализация ──────────────────────────────────────────────────────
  function init() {
    createWidget();
    addLog('Скрипт загружен');

    // запрос разрешения на уведомления
    if (Notification.permission === 'granted') {
      addLog('Разрешение на уведомления уже есть', 'new');
      updateNotifButton();
      startPolling();
    } else if (Notification.permission === 'denied') {
      addLog('Уведомления заблокированы в браузере', 'err');
      setStatus('Уведомления заблокированы', 'error');
      updateNotifButton();
      // всё равно запускаем polling — виджет будет показывать встречи
      startPolling();
    } else {
      setStatus('Нажми «Разрешить уведомления» для старта', '');
      addLog('Ожидаю разрешения на уведомления');
    }
  }

  init();
})();
