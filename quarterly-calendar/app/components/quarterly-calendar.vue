<template>
  <div class="qc-shell">

    <!-- ═══════════════════════════════════
         CALENDAR
    ════════════════════════════════════ -->
    <div class="qc-main">

      <!-- Print-only title (hidden on screen, shown when printing) -->
      <div class="print-title" style="display:none">
        <span class="quarter-label">Q{{ currentQuarter }} {{ currentYear }}</span>
        <span style="font-size:0.9rem;color:#6b6480;font-style:italic">{{ months.map(m => m.name).join(' · ') }}</span>
      </div>

      <!-- Header -->
      <div class="cal-header">
        <button class="nav-btn" @click="prevQuarter">&#8592;</button>
        <div class="quarter-title">
          <span class="quarter-label">Q{{ currentQuarter }}</span>
          <span class="year-label">{{ currentYear }}</span>
          <span class="months-label">{{ months.map(m => m.name.slice(0,3)).join(' · ') }}</span>
        </div>
        <div class="header-actions">
          <button class="nav-btn" @click="nextQuarter">&#8594;</button>
          <button class="pdf-btn" @click="exportPDF" title="Export to PDF">
            <span class="pdf-btn-icon">⬇</span> PDF
          </button>
          <button class="pdf-btn word-btn" @click="exportWord" title="Export to Word (.docx)">
            <span class="pdf-btn-icon">W</span> Word
          </button>
          <button class="pdf-btn term-toggle-btn" :class="{ active: showSchoolTerms }" @click="showSchoolTerms = !showSchoolTerms" title="Toggle school term highlights">
            🏫 Terms
          </button>
          <button class="pdf-btn defaults-btn" @click="loadDefaultEvents" title="Load default recurring events for this quarter">
            ✦ Defaults
          </button>
        </div>
      </div>

      <!-- Calendar Grid -->
      <div class="cal-grid-wrapper">
        <div
          class="cal-grid"
          :style="{ gridTemplateColumns: `56px repeat(${allWeeks.length}, minmax(80px, 1fr))` }"
        >
          <!-- Row 0: Month spans -->
          <div class="corner-cell">Wk →<br/>Day ↓</div>
          <template v-for="(month, mi) in monthsWithCounts" :key="'mh-' + mi">
            <div class="month-span-header" :style="{ gridColumn: `span ${month.weekCount}` }" :class="'mc-' + mi">
              {{ month.name }}
            </div>
          </template>

          <!-- Row 1: Week headers -->
          <div class="corner-sub"></div>
          <div
            v-for="(week, wi) in allWeeks" :key="'wh-' + wi"
            class="week-header"
            :class="{ 'current-week': isCurrentWeek(week), 'month-boundary': week.isMonthBoundary }"
          >
            <span class="week-num">{{ week.weekNumber }}</span>
            <span class="week-date">{{ formatWeekStart(week.startDate) }}</span>
          </div>

          <!-- Day rows -->
          <template v-for="(day, di) in dayNames" :key="'row-' + di">
            <div class="day-label">
              <span class="day-short">{{ day.short }}</span>
            </div>
            <div
              v-for="(week, wi) in allWeeks" :key="'cell-' + di + '-' + wi"
              class="day-cell"
              :class="{
                today: isToday(week.startDate, di),
                'out-of-quarter': !isInQuarter(week.startDate, di),
                'drag-over': dragOverKey === makeCellKey(week.startDate, di),
                'month-boundary': week.isMonthBoundary
              }"
              @click="openAddEvent(week.startDate, di)"
              @dragover.prevent="dragOverKey = makeCellKey(week.startDate, di)"
              @dragleave="dragOverKey = null"
              @drop.prevent="onDropToCell($event, week.startDate, di)"
            >
              <span class="cell-date">{{ getCellDate(week.startDate, di) }}</span>
              <div
                v-if="getSchoolPeriodForCell(week.startDate, di)"
                class="term-band"
                :class="getSchoolPeriodForCell(week.startDate, di).type === 'break' ? 'term-break' : 'term-school'"
                :title="getSchoolPeriodForCell(week.startDate, di).label"
              ></div>
              <div class="events-stack">
                <div
                  v-for="ev in getEventsForCell(week.startDate, di)" :key="ev.id"
                  class="event-chip"
                  :style="{ '--ec': ev.color }"
                  :class="{ 'is-gcal': ev.source === 'gcal' }"
                  :title="ev.title + (ev.time ? ' @ ' + ev.time : '') + '\nClick to remove'"
                  draggable="true"
                  @dragstart="onEventDragStart($event, ev)"
                  @click.stop="removeEvent(ev)"
                >
                  <span v-if="ev.time" class="chip-time">{{ ev.time }}</span>
                  <span class="chip-label">{{ ev.title }}</span>
                  <span v-if="ev.source === 'gcal'" class="chip-gcal-dot" title="From Google Calendar"></span>
                </div>
              </div>
            </div>
          </template>
        </div>
      </div>
    </div>

    <!-- ═══════════════════════════════════
         BOTTOM PANEL
    ════════════════════════════════════ -->
    <div class="bottom-panel">

      <!-- Tab bar -->
      <div class="panel-tabs">
        <button
          v-for="tab in ['palette', 'preacher', 'gcal']" :key="tab"
          class="panel-tab"
          :class="{ active: activeTab === tab }"
          @click="activeTab = tab"
        >
          <span v-if="tab === 'palette'">🎨 Event Types</span>
          <span v-if="tab === 'preacher'">🎤 Guest Preachers</span>
          <span v-if="tab === 'gcal'">📅 Google Calendar</span>
        </button>
        <div class="panel-tab-spacer"></div>
        <span class="panel-hint">Drag events onto the calendar · Click an event to remove it</span>
      </div>

      <!-- ── TAB: Palette ── -->
      <div v-if="activeTab === 'palette'" class="panel-body palette-tab">
        <!-- Existing templates -->
        <div class="palette-chips">
          <div
            v-for="tmpl in eventTemplates" :key="tmpl.id"
            class="palette-chip"
            :style="{ '--c': tmpl.color }"
            draggable="true"
            @dragstart="onPaletteDragStart($event, tmpl)"
          >
            <span class="p-dot"></span>
            <span class="p-label">{{ tmpl.title }}</span>
            <span v-if="tmpl.time" class="p-time">{{ tmpl.time }}</span>
            <span class="p-drag">⠿</span>
            <button class="p-del" @click.stop="deleteTemplate(tmpl)" title="Remove type">×</button>
          </div>
        </div>

        <div class="palette-divider-v"></div>

        <!-- Add new -->
        <div class="palette-new">
          <input v-model="newTmpl.title" class="p-input" placeholder="New type name..." @keyup.enter="addTemplate" />
          <input v-model="newTmpl.time" class="p-input p-time-input" type="time" placeholder="Default time (opt.)" />
          <div class="p-colors">
            <button
              v-for="c in eventColors" :key="c"
              class="p-swatch" :style="{ background: c }"
              :class="{ selected: newTmpl.color === c }"
              @click="newTmpl.color = c"
            ></button>
          </div>
          <button class="p-add-btn" @click="addTemplate" :disabled="!newTmpl.title.trim()">+ Add Type</button>
        </div>
      </div>

      <!-- ── TAB: Guest Preachers ── -->
      <div v-if="activeTab === 'preacher'" class="panel-body preacher-tab">

        <!-- Add form -->
        <div class="preacher-form">
          <div class="preacher-form-title">Add Guest Preacher</div>
          <div class="preacher-form-row">
            <input v-model="preacherNew.name" class="p-input preacher-name-input" placeholder="Preacher name..." @keyup.enter="addPreacher" />
            <div class="session-select">
              <button
                v-for="s in ['AM', 'PM', 'All Day']" :key="s"
                class="session-btn"
                :class="{ selected: preacherNew.session === s }"
                @click="preacherNew.session = s"
              >{{ s }}</button>
            </div>
            <div class="preacher-swatches">
              <button
                v-for="c in preacherColors" :key="c"
                class="p-swatch sm" :style="{ background: c }"
                :class="{ selected: preacherNew.color === c }"
                @click="preacherNew.color = c"
              ></button>
            </div>
            <button class="p-add-btn" @click="addPreacher" :disabled="!preacherNew.name.trim()">+ Add</button>
          </div>
        </div>

        <div class="palette-divider-v"></div>

        <!-- Preacher chips (draggable onto grid) -->
        <div class="palette-chips" style="flex:1">
          <div
            v-for="p in sortedPreachers" :key="p.id"
            class="palette-chip preacher-chip"
            :style="{ '--c': p.color }"
            draggable="true"
            @dragstart="onPaletteDragStart($event, { ...p, title: p.name + (p.session !== 'All Day' ? ' (' + p.session + ')' : '') })"
          >
            <span class="p-dot"></span>
            <span class="p-label">{{ p.name }}</span>
            <span class="preacher-session-badge">{{ p.session }}</span>
            <span class="p-drag">⠿</span>
            <button class="p-del" @click.stop="removePreacher(p)" title="Remove">×</button>
          </div>
          <div v-if="!preachers.length" class="preacher-empty">
            Add a guest preacher above, then drag them onto the calendar
          </div>
        </div>
      </div>

      <!-- ── TAB: Google Calendar ── -->
      <div v-if="activeTab === 'gcal'" class="panel-body gcal-tab">

        <!-- Import section -->
        <div class="gcal-import">
          <div class="gcal-import-header">
            <span class="gcal-icon">📅</span>
            <div>
              <div class="gcal-import-title">Import from Google Calendar</div>
              <div class="gcal-import-sub">Paste ICS/CSV export, or add events manually below</div>
            </div>
          </div>
          <textarea
            v-model="gcalPaste"
            class="gcal-textarea"
            placeholder="Paste Google Calendar .ics content here, or use the manual form below..."
            rows="3"
          ></textarea>
          <button class="gcal-parse-btn" @click="parseICS" :disabled="!gcalPaste.trim()">
            Parse &amp; Import Events
          </button>
          <span v-if="gcalImportMsg" class="gcal-msg" :class="gcalImportOk ? 'ok' : 'err'">{{ gcalImportMsg }}</span>
        </div>

        <div class="gcal-divider"></div>

        <!-- Manual GCal event add -->
        <div class="gcal-manual">
          <div class="gcal-manual-title">Add event manually</div>
          <div class="gcal-form-row">
            <input v-model="gcalNew.title" class="gcal-input" placeholder="Event title" />
            <input v-model="gcalNew.date"  class="gcal-input gcal-date" type="date" />
            <input v-model="gcalNew.time"  class="gcal-input gcal-time" type="time" placeholder="Time (opt.)" />
          </div>
          <div class="gcal-form-row">
            <span class="gcal-assign-label">Assign type:</span>
            <div class="gcal-template-select">
              <button
                v-for="tmpl in eventTemplates" :key="tmpl.id"
                class="gcal-tmpl-btn"
                :style="{ '--c': tmpl.color }"
                :class="{ selected: gcalNew.templateId === tmpl.id }"
                @click="gcalNew.templateId = (gcalNew.templateId === tmpl.id ? null : tmpl.id)"
              >
                <span class="gtb-dot"></span>{{ tmpl.title }}
              </button>
              <button class="gcal-tmpl-btn custom-btn" :class="{ selected: gcalNew.templateId === 'custom' }" @click="gcalNew.templateId = 'custom'" style="--c:#888">
                Custom
              </button>
            </div>
            <div v-if="gcalNew.templateId === 'custom'" class="gcal-custom-color">
              <button
                v-for="c in eventColors" :key="c"
                class="p-swatch sm" :style="{ background: c }"
                :class="{ selected: gcalNew.customColor === c }"
                @click="gcalNew.customColor = c"
              ></button>
            </div>
          </div>
          <button class="gcal-add-btn" @click="addGCalEvent" :disabled="!gcalNew.title.trim() || !gcalNew.date">
            + Add to Calendar
          </button>
        </div>

        <!-- Imported events list -->
        <div v-if="gcalEvents.length" class="gcal-list">
          <div class="gcal-list-title">Imported events ({{ gcalEvents.length }})</div>
          <div class="gcal-list-scroll">
            <div v-for="ev in gcalEvents" :key="ev.id" class="gcal-list-item" :style="{ '--c': ev.color }">
              <span class="gli-dot"></span>
              <span class="gli-title">{{ ev.title }}</span>
              <span class="gli-date">{{ formatGCalDate(ev.date) }}</span>
              <span v-if="ev.time" class="gli-time">{{ ev.time }}</span>
              <span class="gli-source">{{ ev.source === 'gcal' ? '📅' : '✏️' }}</span>
              <!-- Assign template -->
              <div class="gli-assign">
                <select class="gli-select" :value="ev.assignedTemplateId" @change="assignTemplate(ev, $event.target.value)">
                  <option value="">— type —</option>
                  <option v-for="tmpl in eventTemplates" :key="tmpl.id" :value="tmpl.id">{{ tmpl.title }}</option>
                </select>
              </div>
              <button class="gli-del" @click="removeGCalEvent(ev)">×</button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- ═══════════════════════════════════
         ADD EVENT MODAL (click on cell)
    ════════════════════════════════════ -->
    <transition name="modal">
      <div v-if="modal.open" class="modal-backdrop" @click.self="modal.open = false">
        <div class="modal-box">
          <h3 class="modal-title">Add Event</h3>
          <p class="modal-date">{{ formatModalDate(modal.date) }}</p>
          <input v-model="modal.title" class="modal-input" placeholder="Event title..." @keyup.enter="confirmAddEvent" ref="modalInput" />
          <input v-model="modal.time" class="modal-input time-input" type="time" placeholder="Start time (optional)" />
          <!-- Quick assign template -->
          <div class="modal-templates">
            <span class="modal-tmpl-label">Quick type:</span>
            <button
              v-for="tmpl in eventTemplates" :key="tmpl.id"
              class="modal-tmpl-btn"
              :style="{ '--c': tmpl.color }"
              :class="{ selected: modal.templateId === tmpl.id }"
              @click="applyTemplate(tmpl)"
            >
              <span class="mtb-dot"></span>{{ tmpl.title }}
            </button>
          </div>
          <div class="color-picker">
            <button v-for="c in eventColors" :key="c" class="color-swatch" :style="{ background: c }" :class="{ selected: modal.color === c }" @click="modal.color = c"></button>
          </div>
          <div class="modal-actions">
            <button class="btn-cancel" @click="modal.open = false">Cancel</button>
            <button class="btn-confirm" @click="confirmAddEvent" :disabled="!modal.title.trim()">Add Event</button>
          </div>
        </div>
      </div>
    </transition>

  </div>
</template>

<script>
import { zipSync, strToU8 } from 'fflate'
import schoolTermsData from '~/data/schoolTerms.js'

let _uid = 1

export default {
  name: 'QuarterlyCalendar',

  data() {
    const today = new Date()
    return {
      today: new Date(today.getFullYear(), today.getMonth(), today.getDate()),
      currentYear: today.getFullYear(),
      currentQuarter: Math.ceil((today.getMonth() + 1) / 3),
      activeTab: 'palette',

      // Calendar events (placed on grid)
      events: [],

      // GCal imported/manual events (shown in list, can be dragged to grid)
      gcalEvents: [],
      gcalPaste: '',
      gcalImportMsg: '',
      gcalImportOk: false,
      gcalNew: { title: '', date: '', time: '', templateId: null, customColor: '#6366f1' },

      dragOverKey: null,
      draggingTemplate: null,
      draggingEvent: null,
      draggingGCal: null,

      eventColors: ['#6366f1','#ec4899','#f59e0b','#10b981','#3b82f6','#f43f5e','#8b5cf6','#14b8a6'],

      eventTemplates: [
        { id: _uid++, title: 'Prayer Meeting',    time: '19:30', color: '#6366f1' },
        { id: _uid++, title: 'Discoverers',       time: '19:30', color: '#f43f5e' },
        { id: _uid++, title: 'First Steps',       time: '9:00',  color: '#ffeffe' },
        { id: _uid++, title: 'Coffee Morning',    time: '10:00', color: '#f59e0b' },
        { id: _uid++, title: 'Small Groups',      time: '10:00', color: '#10b981' },
        { id: _uid++, title: 'Small Groups',      time: '19:30', color: '#10b981' },
        { id: _uid++, title: 'Small Groups',      time: '19:45', color: '#10b981' },
        { id: _uid++, title: 'Members Meeting',   time: '',      color: '#ec4899' },
        { id: _uid++, title: "Men's Meeting",     time: '',      color: '#8b5cf6' },
        { id: _uid++, title: 'Ladies Meeting',    time: '',      color: '#14b8a6' },
        { id: _uid++, title: '(Rooted and Grounded)', time: '9:00', color: '#14b8a6' },
        { id: _uid++, title: 'Joint Meeting',     time: '',      color: '#14b8a6' },
        { id: _uid++, title: 'Pastor',            time: '',      color: '#60a5fa' },
      ],

      newTmpl: { title: '', time: '', color: '#6366f1' },

      // School terms loaded from ~/data/schoolTerms.js — edit that file to update
      schoolPeriods: schoolTermsData,
      showSchoolTerms: true,

      // Guest Preacher entries — sorted: All Day → AM → PM
      preachers: [],
      preacherNew: { name: '', session: 'AM', color: '#60a5fa' },
      preacherColors: ['#60a5fa','#34d399','#f59e0b','#f43f5e','#a78bfa','#ec4899','#14b8a6','#fb923c'],

      modal: { open: false, date: null, title: '', time: '', color: '#6366f1', templateId: null },

      dayNames: [
        { short: 'Sun' },
        { short: 'Mon' }, { short: 'Tue' }, { short: 'Wed' },
        { short: 'Thu' }, { short: 'Fri' }, { short: 'Sat' }
      ]
    }
  },

  computed: {
    quarterStartMonth() { return (this.currentQuarter - 1) * 3 },

    months() {
      const names = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December']
      return [0,1,2].map(i => ({ name: names[this.quarterStartMonth + i], index: this.quarterStartMonth + i }))
    },

    allWeeks() {
      const year = this.currentYear
      const qStart = this.quarterStartMonth
      const firstOfQ = new Date(year, qStart, 1)
      const lastOfQ  = new Date(year, qStart + 3, 0)
      let cur = new Date(firstOfQ)
      const dow = cur.getDay()
      cur.setDate(cur.getDate() - dow)   // snap back to Sunday
      const weeks = []
      let prevRepMonth = -1
      while (cur <= lastOfQ) {
        const wStart = new Date(cur)
        const wed = new Date(wStart); wed.setDate(wed.getDate() + 3)
        const repMonth = wed.getMonth()
        const isMonthBoundary = repMonth !== prevRepMonth && prevRepMonth !== -1
        prevRepMonth = repMonth
        weeks.push({ startDate: wStart, weekNumber: this.getISOWeek(wStart), isMonthBoundary, repMonth })
        cur.setDate(cur.getDate() + 7)
      }
      return weeks
    },

    sortedPreachers() {
      const order = { 'All Day': 0, 'AM': 1, 'PM': 2 }
      return [...this.preachers].sort((a, b) => (order[a.session] ?? 9) - (order[b.session] ?? 9))
    },

    activeSchoolPeriods() {
      if (!this.showSchoolTerms) return []
      const qStart = new Date(this.currentYear, this.quarterStartMonth, 1)
      const qEnd   = new Date(this.currentYear, this.quarterStartMonth + 3, 0)
      return this.schoolPeriods.map(p => {
        const s = new Date(p.start), e = new Date(p.end)
        if (e < qStart || s > qEnd) return null
        return { ...p, startD: s, endD: e }
      }).filter(Boolean)
    },

    monthsWithCounts() {
      const counts = {}
      this.months.forEach(m => { counts[m.index] = 0 })
      this.allWeeks.forEach(w => {
        if (counts[w.repMonth] !== undefined) counts[w.repMonth]++
        else {
          const closest = this.months.reduce((b, m) => Math.abs(m.index - w.repMonth) < Math.abs(b.index - w.repMonth) ? m : b)
          counts[closest.index]++
        }
      })
      return this.months.map(m => ({ ...m, weekCount: counts[m.index] || 1 }))
    }
  },

  methods: {
    // ── Navigation ──
    prevQuarter() { this.currentQuarter--; if (this.currentQuarter < 1) { this.currentQuarter = 4; this.currentYear-- } },
    nextQuarter() { this.currentQuarter++; if (this.currentQuarter > 4) { this.currentQuarter = 1; this.currentYear++ } },

    // ── Cell helpers ──
    getCellDate(weekStart, di) { const d = new Date(weekStart); d.setDate(d.getDate() + di); return d.getDate() },
    getCellDateObj(weekStart, di) { const d = new Date(weekStart); d.setDate(d.getDate() + di); return d },
    makeCellKey(weekStart, di) {
      const d = this.getCellDateObj(weekStart, di)
      return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`
    },
    dateToKey(date) { return `${date.getFullYear()}-${date.getMonth()}-${date.getDate()}` },
    isToday(weekStart, di) { return this.getCellDateObj(weekStart, di).getTime() === this.today.getTime() },
    isInQuarter(weekStart, di) {
      const m = this.getCellDateObj(weekStart, di).getMonth()
      return this.months.some(mo => mo.index === m)
    },
    isCurrentWeek(week) {
      const now = new Date(); const end = new Date(week.startDate); end.setDate(end.getDate() + 6)
      return now >= week.startDate && now <= end
    },
    formatWeekStart(date) { return date.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) },
    formatModalDate(date) {
      if (!date) return ''
      return date.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })
    },
    formatGCalDate(dateStr) {
      if (!dateStr) return ''
      const d = new Date(dateStr + 'T00:00:00')
      return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
    },
    getSchoolPeriodForCell(weekStart, di) {
      if (!this.showSchoolTerms) return null
      const d = this.getCellDateObj(weekStart, di)
      const t = d.getTime()
      return this.activeSchoolPeriods.find(p => {
        const s = new Date(p.startD); s.setHours(0,0,0,0)
        const e = new Date(p.endD); e.setHours(23,59,59,999)
        return t >= s.getTime() && t <= e.getTime()
      }) || null
    },

    getEventsForCell(weekStart, di) {
      const key = this.makeCellKey(weekStart, di)
      return this.events.filter(e => e.key === key).sort((a,b) => { const ta = a.time || '99:99', tb = b.time || '99:99'; return ta.localeCompare(tb) })
    },

    // ── Drag: palette template ──
    onPaletteDragStart(e, tmpl) {
      this.draggingTemplate = tmpl; this.draggingEvent = null; this.draggingGCal = null
      e.dataTransfer.effectAllowed = 'copy'
    },
    // ── Drag: existing placed event ──
    onEventDragStart(e, ev) {
      this.draggingEvent = ev; this.draggingTemplate = null; this.draggingGCal = null
      e.dataTransfer.effectAllowed = 'move'; e.stopPropagation()
    },

    // ── Drop onto cell ──
    onDropToCell(e, weekStart, di) {
      this.dragOverKey = null
      const d = this.getCellDateObj(weekStart, di)
      const key = this.dateToKey(d)

      if (this.draggingTemplate) {
        this.events.push({ id: _uid++, key, title: this.draggingTemplate.title, color: this.draggingTemplate.color, time: this.draggingTemplate.time || '', source: 'manual', date: new Date(d) })
        this.draggingTemplate = null
      } else if (this.draggingEvent) {
        this.draggingEvent.key = key; this.draggingEvent.date = new Date(d)
        this.draggingEvent = null
      }
    },

    // ── Modal (click cell) ──
    openAddEvent(weekStart, di) {
      const d = this.getCellDateObj(weekStart, di)
      this.modal = { open: true, date: d, title: '', time: '', color: this.eventColors[0], templateId: null }
      this.$nextTick(() => this.$refs.modalInput?.focus())
    },
    applyTemplate(tmpl) {
      this.modal.templateId = tmpl.id; this.modal.title = this.modal.title || tmpl.title; this.modal.color = tmpl.color
    },
    confirmAddEvent() {
      if (!this.modal.title.trim()) return
      const d = this.modal.date
      this.events.push({
        id: _uid++,
        key: this.dateToKey(d),
        title: this.modal.title.trim(),
        time: this.modal.time || '',
        color: this.modal.color,
        source: 'manual',
        date: new Date(d)
      })
      this.modal.open = false
    },
    removeEvent(ev) {
      const idx = this.events.indexOf(ev); if (idx > -1) this.events.splice(idx, 1)
    },

    // ── Guest Preachers ──
    addPreacher() {
      if (!this.preacherNew.name.trim()) return
      this.preachers.push({
        id: _uid++,
        name: this.preacherNew.name.trim(),
        session: this.preacherNew.session,
        color: this.preacherNew.color,
        time: '',
      })
      this.preacherNew = { name: '', session: 'AM', color: this.preacherColors[0] }
    },
    removePreacher(p) {
      const i = this.preachers.indexOf(p); if (i > -1) this.preachers.splice(i, 1)
    },

    // ── Templates ──
    addTemplate() {
      if (!this.newTmpl.title.trim()) return
      this.eventTemplates.push({ id: _uid++, title: this.newTmpl.title.trim(), time: this.newTmpl.time || '', color: this.newTmpl.color })
      this.newTmpl = { title: '', time: '', color: this.eventColors[0] }
    },
    deleteTemplate(tmpl) {
      const idx = this.eventTemplates.indexOf(tmpl); if (idx > -1) this.eventTemplates.splice(idx, 1)
    },

    // ── GCal: manual add ──
    addGCalEvent() {
      if (!this.gcalNew.title.trim() || !this.gcalNew.date) return
      const tmpl = this.gcalNew.templateId && this.gcalNew.templateId !== 'custom'
        ? this.eventTemplates.find(t => t.id === this.gcalNew.templateId)
        : null
      const color = tmpl ? tmpl.color : (this.gcalNew.customColor || '#888')
      const ev = {
        id: _uid++,
        title: this.gcalNew.title.trim(),
        date: this.gcalNew.date,
        time: this.gcalNew.time || '',
        color,
        source: 'manual-gcal',
        assignedTemplateId: tmpl ? tmpl.id : null
      }
      this.gcalEvents.push(ev)
      // Also place on calendar grid
      this.placeGCalEventOnGrid(ev)
      this.gcalNew = { title: '', date: '', time: '', templateId: null, customColor: '#6366f1' }
    },

    placeGCalEventOnGrid(ev) {
      if (!ev.date) return
      const d = new Date(ev.date + 'T00:00:00')
      const key = this.dateToKey(d)
      // avoid duplicates
      const exists = this.events.find(e => e.gcalId === ev.id)
      if (exists) { exists.key = key; exists.time = ev.time; exists.color = ev.color; return }
      this.events.push({
        id: _uid++,
        gcalId: ev.id,
        key,
        title: ev.title,
        time: ev.time || '',
        color: ev.color,
        source: 'gcal',
        date: d
      })
    },

    assignTemplate(ev, templateId) {
      ev.assignedTemplateId = templateId || null
      const tmpl = templateId ? this.eventTemplates.find(t => String(t.id) === String(templateId)) : null
      if (tmpl) {
        ev.color = tmpl.color
        // Update the placed grid event too
        const placed = this.events.find(e => e.gcalId === ev.id)
        if (placed) placed.color = tmpl.color
      }
    },

    removeGCalEvent(ev) {
      const idx = this.gcalEvents.indexOf(ev); if (idx > -1) this.gcalEvents.splice(idx, 1)
      // Also remove from grid
      const pi = this.events.findIndex(e => e.gcalId === ev.id)
      if (pi > -1) this.events.splice(pi, 1)
    },

    // ── GCal ICS parser ──
    parseICS() {
      this.gcalImportMsg = ''; this.gcalImportOk = false
      const text = this.gcalPaste
      const events = []
      const veventBlocks = text.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || []

      if (veventBlocks.length === 0) {
        this.gcalImportMsg = 'No events found. Make sure you pasted valid .ics content.'; return
      }

      veventBlocks.forEach(block => {
        const get = (key) => {
          const m = block.match(new RegExp(key + '[^:]*:(.+)'))
          return m ? m[1].replace(/\r/g, '').trim() : ''
        }
        const summary = get('SUMMARY')
        const dtstart = get('DTSTART')
        if (!summary || !dtstart) return

        // Parse date from DTSTART:20240315 or DTSTART:20240315T090000Z
        const dateStr = dtstart.replace(/T.*/, '').replace(/[^0-9]/g,'')
        if (dateStr.length < 8) return
        const year = dateStr.slice(0,4), month = dateStr.slice(4,6), day = dateStr.slice(6,8)
        const isoDate = `${year}-${month}-${day}`
        let time = ''
        const timePart = dtstart.match(/T(\d{2})(\d{2})/)
        if (timePart) {
          // Note: ICS times are UTC; for simplicity display as-is
          time = `${timePart[1]}:${timePart[2]}`
        }

        events.push({
          id: _uid++,
          title: summary,
          date: isoDate,
          time,
          color: this.eventColors[0],
          source: 'gcal',
          assignedTemplateId: null
        })
      })

      if (events.length === 0) {
        this.gcalImportMsg = 'Could not parse any events. Check the format.'; return
      }

      events.forEach(ev => {
        this.gcalEvents.push(ev)
        this.placeGCalEventOnGrid(ev)
      })

      this.gcalPaste = ''
      this.gcalImportMsg = `✓ Imported ${events.length} event${events.length > 1 ? 's' : ''}`
      this.gcalImportOk = true
    },

    // ── Word Export (fflate — no docx dependency) ──
    exportWord() {
      const weeks = this.allWeeks
      const months = this.monthsWithCounts
      const quarterTitle = `Q${this.currentQuarter} ${this.currentYear} — ${this.months.map(m => m.name).join(', ')}`
      const printDate = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })

      // ── XML helpers ──
      const esc = (s) => String(s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&apos;')

      const hex = (color) => color.replace('#', '').toUpperCase()

      // A4 landscape DXA. 1 inch = 1440 DXA. 8mm margin ≈ 454 DXA.
      const MARGIN = 454
      const CONTENT_W = 16838 - MARGIN * 2   // 15930 DXA available
      const DAY_COL   = 560
      const weekColW  = Math.floor((CONTENT_W - DAY_COL) / weeks.length)
      const lastColW  = CONTENT_W - DAY_COL - weekColW * (weeks.length - 1)

      const colW = (wi) => wi === weeks.length - 1 ? lastColW : weekColW

      // ── Cell XML builders ──
      const borderXml = (leftThick) => `
        <w:tcBorders>
          <w:top    w:val="single" w:sz="4"  w:color="AAAAAA"/>
          <w:bottom w:val="single" w:sz="4"  w:color="AAAAAA"/>
          <w:left   w:val="single" w:sz="${leftThick ? 12 : 4}" w:color="${leftThick ? '666666' : 'AAAAAA'}"/>
          <w:right  w:val="single" w:sz="4"  w:color="AAAAAA"/>
        </w:tcBorders>`

      const cellPr = (width, fill, leftThick, vAlign, span) => `
        <w:tcPr>
          <w:tcW w:w="${width}" w:type="dxa"/>
          ${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}
          ${fill ? `<w:shd w:val="clear" w:fill="${fill}"/>` : ''}
          ${borderXml(leftThick)}
          <w:tcMar>
            <w:top    w:w="40"  w:type="dxa"/>
            <w:bottom w:w="40"  w:type="dxa"/>
            <w:left   w:w="80"  w:type="dxa"/>
            <w:right  w:w="80"  w:type="dxa"/>
          </w:tcMar>
          ${vAlign ? `<w:vAlign w:val="${vAlign}"/>` : ''}
        </w:tcPr>`

      const para = (text, opts = {}) => {
        const jc    = opts.align  ? `<w:jc w:val="${opts.align}"/>` : ''
        const ind   = opts.indent ? `<w:ind w:left="${opts.indent}"/>` : ''
        const pBdr  = opts.leftColor
          ? `<w:pBdr><w:left w:val="single" w:sz="8" w:space="2" w:color="${hex(opts.leftColor)}"/></w:pBdr>`
          : ''
        const sz    = opts.size   || 18
        const bold  = opts.bold   ? '<w:b/>' : ''
        const color = opts.color  ? `<w:color w:val="${opts.color}"/>` : ''
        const runs  = (opts.runs || [{ text, ...opts }]).map(r =>
          `<w:r><w:rPr><w:sz w:val="${r.size || sz}"/><w:szCs w:val="${r.size || sz}"/>${r.bold ? '<w:b/>' : ''}${r.color ? `<w:color w:val="${r.color}"/>` : ''}</w:rPr><w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`
        ).join('')
        return `<w:p><w:pPr>${jc}${ind}${pBdr}</w:pPr>${runs}</w:p>`
      }

      const cell = (content, width, fill, leftThick = false, vAlign = '', span = 1) =>
        `<w:tc>${cellPr(width, fill, leftThick, vAlign, span)}${content || '<w:p/>'}</w:tc>`

      // ── Row 0: Month headers ──
      const monthFills  = ['EDE9FE', 'DBEAFE', 'D1FAE5']
      const monthColors = ['5B21B6', '1D4ED8', '065F46']
      let monthHeaderRow = `<w:tr><w:trPr><w:tblHeader/></w:trPr>`
      monthHeaderRow += cell('', DAY_COL, 'F0F0F0')
      months.forEach((m, mi) => {
        const spanCount = m.weekCount
        const width = weekColW * spanCount + (mi === months.length - 1 ? lastColW - weekColW : 0)
        monthHeaderRow += cell(
          para(m.name.toUpperCase(), { align: 'center', bold: true, size: 16, color: monthColors[mi] }),
          width, monthFills[mi], mi > 0, '', spanCount
        )
      })
      monthHeaderRow += '</w:tr>'

      // ── Row 1: Week headers ──
      let weekHeaderRow = `<w:tr><w:trPr><w:tblHeader/></w:trPr>`
      weekHeaderRow += cell('', DAY_COL, 'F0F0F0')
      weeks.forEach((w, wi) => {
        const label = w.startDate.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })
        weekHeaderRow += cell(
          para('', { align: 'center', runs: [
            { text: `Wk ${w.weekNumber}`, size: 14, bold: true, color: '444444' }
          ]}) +
          para('', { align: 'center', runs: [
            { text: label, size: 12, color: '888888' }
          ]}),
          colW(wi), 'F8F8F8', w.isMonthBoundary
        )
      })
      weekHeaderRow += '</w:tr>'

      // ── Day rows ──
      const DAY_NAMES = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
      const dayRows = DAY_NAMES.map((dayName, di) => {
        let row = `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="0" w:hRule="auto"/></w:trPr>`

        // Day label cell
        row += cell(
          para(dayName, { align: 'center', bold: true, size: 16, color: '333333' }),
          DAY_COL, 'F8F8F8', false, 'center'
        )

        weeks.forEach((week, wi) => {
          const d = new Date(week.startDate); d.setDate(d.getDate() + di)
          const inQ  = this.months.some(mo => mo.index === d.getMonth())
          const key  = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`
          const evs  = inQ
            ? this.events.filter(e => e.key === key).sort((a,b) => (a.time||'').localeCompare(b.time||''))
            : []

          let cellContent = ''
          if (inQ) {
            cellContent += para(String(d.getDate()), { size: 12, color: 'BBBBBB' })
            evs.forEach(ev => {
              cellContent += para('', {
                leftColor: ev.color,
                indent: 80,
                runs: ev.time
                  ? [{ text: ev.time + ' ', size: 14, color: '666666' }, { text: ev.title, size: 14, bold: true, color: '111111' }]
                  : [{ text: ev.title, size: 14, color: '222222' }]
              })
            })
          }

          row += cell(
            cellContent || '<w:p/>',
            colW(wi),
            inQ ? 'FFFFFF' : 'F5F5F5',
            week.isMonthBoundary
          )
        })

        row += '</w:tr>'
        return row
      }).join('')

      // ── Grid column defs ──
      const gridCols = [DAY_COL, ...weeks.map((_,wi) => colW(wi))]
        .map(w => `<w:gridCol w:w="${w}"/>`)
        .join('')

      // ── Full document XML ──
      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p>
      <w:pPr><w:spacing w:after="80"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>${esc(quarterTitle)}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="120"/></w:pPr>
      <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="888888"/></w:rPr><w:t>Printed ${esc(printDate)}</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="${CONTENT_W}" w:type="dxa"/>
        <w:tblBorders>
          <w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/>
          <w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/>
        </w:tblBorders>
        <w:tblLook w:val="0000"/>
      </w:tblPr>
      <w:tblGrid>${gridCols}</w:tblGrid>
      ${monthHeaderRow}
      ${weekHeaderRow}
      ${dayRows}
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
      <w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="0" w:footer="0"/>
    </w:sectPr>
  </w:body>
</w:document>`

      // ── Required .docx support files ──
      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

      const wordRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`

      const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`

      const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`

      // ── Zip with fflate and download ──
      const zipped = zipSync({
        '[Content_Types].xml':    strToU8(contentTypesXml),
        '_rels/.rels':            strToU8(relsXml),
        'word/document.xml':      strToU8(docXml),
        'word/_rels/document.xml.rels': strToU8(wordRelsXml),
        'word/settings.xml':      strToU8(settingsXml),
      }, { level: 6 })

      const blob = new Blob([zipped], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
      const url  = URL.createObjectURL(blob)
      const a    = document.createElement('a')
      a.href     = url
      a.download = `Q${this.currentQuarter}-${this.currentYear}-calendar.docx`
      a.click()
      URL.revokeObjectURL(url)
    },

    // ── PDF Export ──
    exportPDF() {
      // Build a completely separate minimal HTML document in a new window
      const weeks = this.allWeeks
      const dayNames = this.dayNames
      const monthsInfo = this.monthsWithCounts

      // Helper: get events for a specific date key
      const getEvents = (weekStart, di) => {
        const d = new Date(weekStart); d.setDate(d.getDate() + di)
        const key = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`
        return this.events
          .filter(e => e.key === key)
          .sort((a, b) => (a.time || '').localeCompare(b.time || ''))
      }

      const getCellDate = (weekStart, di) => {
        const d = new Date(weekStart); d.setDate(d.getDate() + di); return d.getDate()
      }

      const isInQuarter = (weekStart, di) => {
        const d = new Date(weekStart); d.setDate(d.getDate() + di)
        return this.months.some(mo => mo.index === d.getMonth())
      }

      // ── Build table HTML ──
      const colCount = weeks.length

      // Month header row
      let monthHeaders = '<th style="width:36px;border:1px solid #999;padding:3px 4px;font-size:7pt;text-align:center;background:#f0f0f0;"></th>'
      monthsInfo.forEach(m => {
        monthHeaders += `<th colspan="${m.weekCount}" style="border:1px solid #999;padding:4px 6px;font-size:8pt;font-weight:bold;text-align:center;background:#e8e8e8;letter-spacing:0.05em;">${m.name.toUpperCase()}</th>`
      })

      // Week header row
      let weekHeaders = '<th style="border:1px solid #999;padding:3px 2px;font-size:6pt;text-align:center;background:#f0f0f0;"></th>'
      weeks.forEach(w => {
        const d = w.startDate
        const label = d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })
        const boundary = w.isMonthBoundary ? 'border-left:2px solid #555;' : ''
        weekHeaders += `<th style="border:1px solid #bbb;${boundary}padding:3px 2px;font-size:6pt;font-weight:600;text-align:center;background:#f8f8f8;min-width:0;"><div style="font-size:6.5pt;color:#555;">Wk ${w.weekNumber}</div><div style="font-size:5.5pt;color:#888;">${label}</div></th>`
      })

      // Day rows
      let dayRows = ''
      dayNames.forEach((day, di) => {
        let cells = `<td style="border:1px solid #999;padding:3px 4px;font-size:7pt;font-weight:600;text-align:center;background:#fff;color:#333;white-space:nowrap;">${day.short}</td>`

        weeks.forEach(week => {
          const inQ = isInQuarter(week.startDate, di)
          const dateNum = getCellDate(week.startDate, di)
          const evList = getEvents(week.startDate, di)
          const boundary = week.isMonthBoundary ? 'border-left:2px solid #555;' : ''
          const cellBg = !inQ ? 'background:#f5f5f5;' : 'background:#fff;'

          let evHTML = ''
          evList.forEach(ev => {
            const timeStr = ev.time ? `<span style="font-size:5pt;color:#555;margin-right:2px;">${ev.time}</span>` : ''
            evHTML += `<div style="font-size:5.5pt;line-height:1.3;margin-top:1px;padding:1px 2px;border-left:2px solid ${ev.color};padding-left:3px;">${timeStr}<span style="color:#111;">${ev.title}</span></div>`
          })

          cells += `<td style="border:1px solid #ddd;${boundary}${cellBg}padding:2px 3px;vertical-align:top;min-width:0;">
            ${inQ ? `<div style="font-size:5.5pt;color:#aaa;line-height:1;margin-bottom:1px;">${dateNum}</div>${evHTML}` : ''}
          </td>`
        })

        dayRows += `<tr>${cells}</tr>`
      })

      const quarterTitle = `Q${this.currentQuarter} ${this.currentYear} — ${this.months.map(m => m.name).join(', ')}`

      const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>${quarterTitle}</title>
  <style>
    @page { size: A4 landscape; margin: 8mm; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html, body {
      height: 100%;
      font-family: Arial, Helvetica, sans-serif;
      font-size: 8pt;
      color: #000;
      background: #fff;
    }
    .page {
      display: flex;
      flex-direction: column;
      height: 100%;
    }
    h1 {
      font-size: 11pt;
      font-weight: bold;
      margin-bottom: 2px;
      color: #111;
      letter-spacing: 0.02em;
      flex-shrink: 0;
    }
    .sub {
      font-size: 7pt;
      color: #666;
      margin-bottom: 6px;
      flex-shrink: 0;
    }
    .table-wrap {
      flex: 1;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }
    table {
      width: 100%;
      height: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }
    thead tr th {
      flex-shrink: 0;
    }
    tbody {
      height: 100%;
    }
    tbody tr {
      /* 7 day rows share all remaining height equally */
      height: calc(100% / 7);
    }
    td, th {
      overflow: hidden;
      word-break: break-word;
      vertical-align: top;
    }
    @media print {
      html, body { height: 100%; margin: 0; }
      h1 { font-size: 10pt; }
    }
  </style>
</head>
<body>
  <div class="page">
    <h1>${quarterTitle}</h1>
    <div class="sub">Printed ${new Date().toLocaleDateString('en-GB', { weekday:'long', day:'numeric', month:'long', year:'numeric' })}</div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>${monthHeaders}</tr>
          <tr>${weekHeaders}</tr>
        </thead>
        <tbody>
          ${dayRows}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>`

      const win = window.open('', '_blank', 'width=1200,height=800')
      if (!win) { alert('Please allow pop-ups to export PDF'); return }
      win.document.write(html)
      win.document.close()
      win.focus()
      setTimeout(() => { win.print() }, 400)
    },

    // ── Load Default Events ──
    loadDefaultEvents() {
      const year  = this.currentYear
      const qStart = this.quarterStartMonth

      // Helper: push a grid event
      const add = (date, title, time, color) => {
        const key = this.dateToKey(date)
        // avoid exact duplicates (same title + time + key)
        const dupe = this.events.find(e => e.key === key && e.title === title && e.time === time)
        if (dupe) return
        this.events.push({ id: _uid++, key, title, time, color, source: 'manual', date: new Date(date) })
      }

      // Walk every day of the quarter
      for (let m = 0; m < 3; m++) {
        const monthIdx = qStart + m
        const daysInMonth = new Date(year, monthIdx + 1, 0).getDate()

        // Find the last Wednesday in this month (for Small Groups week detection)
        let lastWed = new Date(year, monthIdx, daysInMonth)
        while (lastWed.getDay() !== 3) lastWed.setDate(lastWed.getDate() - 1)
        const lastWedKey = this.dateToKey(lastWed)

        // Find the last Thursday in this month (Small Groups Thu)
        let lastThu = new Date(year, monthIdx, daysInMonth)
        while (lastThu.getDay() !== 4) lastThu.setDate(lastThu.getDate() - 1)

        for (let d = 1; d <= daysInMonth; d++) {
          const date = new Date(year, monthIdx, d)
          const dow  = date.getDay() // 0=Sun, 3=Wed, 4=Thu

          // ── Sunday: Pastor Paul ──
          if (dow === 0) {
            add(date, 'Pastor Paul', '', '#60a5fa')
          }

          // ── Wednesday ──
          if (dow === 3) {
            const key = this.dateToKey(date)
            if (key === lastWedKey) {
              // Small Groups week — skip Bible Study/PM, add 3 Small Groups events
              add(date, 'Small Groups', '10:00', '#10b981')
              add(date, 'Small Groups', '19:30', '#10b981')
            } else {
              // Regular week — Bible Study & PM (combined)
              add(date, 'Bible Study & PM', '19:30', '#6366f1')
            }
          }

          // ── Thursday: Small Groups 19:45 on last-week Thursday ──
          if (dow === 4) {
            // same week as lastWed = within 3 days after lastWed
            const diff = date.getTime() - lastWed.getTime()
            if (diff >= 0 && diff <= 3 * 86400000) {
              add(date, 'Small Groups', '19:45', '#10b981')
            }
          }
        }
      }
    },

    getISOWeek(date) {
      const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()))
      const day = d.getUTCDay() || 7; d.setUTCDate(d.getUTCDate() + 4 - day)
      const y = new Date(Date.UTC(d.getUTCFullYear(), 0, 1))
      return Math.ceil((((d - y) / 86400000) + 1) / 7)
    }
  }
}
</script>

<style scoped>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,300&family=DM+Mono:wght@400;500&display=swap');
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

/* ══════════════════════════════
   SHELL
══════════════════════════════ */
.qc-shell {
  display: flex;
  flex-direction: column;
  font-family: 'DM Sans', sans-serif;
  background: #f8fafc;
  color: #1e293b;
  border-radius: 18px;
  overflow: hidden;
  border: 1px solid #e2e8f0;
}

/* ══════════════════════════════
   CALENDAR AREA
══════════════════════════════ */
.qc-main {
  flex: 1;
  display: flex;
  flex-direction: column;
  padding: 18px 16px 12px;
  background: #f8fafc;
}

.cal-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 14px;
}
.quarter-title { display: flex; align-items: baseline; gap: 10px; }
.quarter-label { font-family: 'DM Mono', monospace; font-size: 1.7rem; font-weight: 500; color: #334155; letter-spacing: -1px; }
.year-label { font-size: 1rem; color: #1e293b; font-weight: 300; }
.months-label { font-size: 0.72rem; color: #1e293b; font-style: italic; }
.nav-btn {
  background: #fff; border: 1px solid #e2e8f0; color: #334155;
  width: 30px; height: 30px; border-radius: 8px; cursor: pointer; font-size: 0.9rem;
  transition: background 0.1s, border-color 0.1s, color 0.1s;
}
.nav-btn:hover { background: #f1f5f9; border-color: #1e293b; color: #1e293b; }

.header-actions { display: flex; align-items: center; gap: 8px; }
.pdf-btn {
  display: flex; align-items: center; gap: 5px;
  background: #fff; border: 1px solid #e2e8f0; color: #334155;
  padding: 0 13px; height: 30px; border-radius: 8px; cursor: pointer;
  font-family: 'DM Sans', sans-serif; font-size: 0.75rem; font-weight: 500;
  transition: background 0.1s, border-color 0.1s, color 0.1s; white-space: nowrap;
}
.pdf-btn:hover { background: #f1f5f9; border-color: #1e293b; color: #1e293b; }
.pdf-btn-icon { font-size: 0.7rem; }
.word-btn { border-color: #bbf7d0; color: #16a34a; }
.word-btn:hover { background: #f0fdf4; border-color: #4ade80; color: #15803d; }
.term-toggle-btn { border-color: #d1fae5; color: #059669; }
.term-toggle-btn:hover { background: #ecfdf5; border-color: #34d399; color: #047857; }
.term-toggle-btn.active { background: #ecfdf5; border-color: #10b981; color: #065f46; }
.defaults-btn { border-color: #fde68a; color: #92400e; }
.defaults-btn:hover { background: #fffbeb; border-color: #f59e0b; color: #78350f; }

.cal-grid-wrapper { overflow-x: auto; }
.cal-grid {
  display: grid; gap: 1px;
  background: #e2e8f0; border: 1px solid #e2e8f0; border-radius: 12px;
  overflow: hidden; width: 100%;
}

.corner-cell {
  background: #f1f5f9; padding: 6px 5px 4px;
  font-size: 0.48rem; color: #1e293b; text-transform: uppercase; letter-spacing: 0.06em;
  display: flex; align-items: flex-end; line-height: 1.5;
}
.corner-sub { background: #f1f5f9; }

.month-span-header {
  padding: 6px 8px 4px; font-size: 0.65rem; font-weight: 600;
  letter-spacing: 0.08em; text-transform: uppercase; text-align: center;
}
.mc-0, .mc-1, .mc-2 { background: #f1f5f9; color: #1e293b; border-bottom: 1px solid #e2e8f0; }

.week-header { background: #f1f5f9; padding: 5px 3px; text-align: center; }
.week-header.current-week { background: #eff6ff; }
.week-header.month-boundary { border-left: 2px solid #e2e8f0; }
.week-num { display: block; font-family: 'DM Mono', monospace; font-size: 0.58rem; color: #1e293b; line-height: 1; margin-bottom: 2px; }
.week-date { display: block; font-size: 0.54rem; color: #1e293b; white-space: nowrap; }

.day-label { background: #f1f5f9; padding: 0 5px; display: flex; align-items: center; justify-content: center; min-height: 54px; }
.day-short { font-family: 'DM Mono', monospace; font-size: 0.6rem; color: #1e293b; font-weight: 500; }

.day-cell {
  background: #fff; min-height: 54px; padding: 3px 3px 2px;
  position: relative; cursor: pointer; transition: background 0.08s; overflow: hidden;
}
.day-cell.month-boundary { border-left: 2px solid #e2e8f0; }
.day-cell:hover { background: #f8fafc; }
.day-cell.out-of-quarter { opacity: 0.25; pointer-events: none; }
.day-cell.today { background: #eff6ff !important; }
.day-cell.today::after { content: ''; position: absolute; inset: 0; border: 1px solid #bfdbfe; pointer-events: none; border-radius: 1px; }
.day-cell.drag-over { background: #dbeafe !important; box-shadow: inset 0 0 0 2px #3b82f6; }

.cell-date { font-family: 'DM Mono', monospace; font-size: 0.56rem; color: #1e293b; display: block; margin-bottom: 2px; line-height: 1; }
.day-cell.today .cell-date { color: #3b82f6; font-weight: 600; }

.events-stack { display: flex; flex-direction: column; gap: 1px; }
.event-chip {
  display: flex; align-items: center; gap: 2px;
  padding: 2px 4px; border-radius: 3px;
  background: color-mix(in srgb, var(--ec) 18%, #fff);
  border-left: 2px solid var(--ec);
  cursor: grab; transition: opacity 0.1s, transform 0.1s; max-width: 100%;
}
.event-chip:hover { opacity: 0.75; transform: scale(1.03); }
.event-chip:active { cursor: grabbing; }
.event-chip.is-gcal { border-style: dashed; }

.chip-time {
  font-family: 'DM Mono', monospace;
  font-size: 0.5rem; color: rgba(0,0,0,0.45);
  background: rgba(0,0,0,0.07);
  border-radius: 3px; padding: 0 3px;
  flex-shrink: 0; letter-spacing: -0.01em; font-weight: 600;
}
.chip-label {
  font-size: 0.56rem; font-weight: 600; color: #1e293b;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis; pointer-events: none; flex: 1;
}
.chip-gcal-dot {
  width: 4px; height: 4px; border-radius: 50%;
  background: rgba(0,0,0,0.25); flex-shrink: 0;
}

/* ══════════════════════════════
   BOTTOM PANEL
══════════════════════════════ */
.bottom-panel {
  border-top: 1px solid #e2e8f0;
  background: #f8fafc;
  flex-shrink: 0;
}

.panel-tabs {
  display: flex; align-items: center; gap: 0;
  border-bottom: 1px solid #e2e8f0; padding: 0 16px;
}
.panel-tab {
  background: transparent; border: none; border-bottom: 2px solid transparent;
  color: #1e293b; padding: 10px 16px; cursor: pointer;
  font-family: 'DM Sans', sans-serif; font-size: 0.78rem; font-weight: 500;
  transition: color 0.13s, border-color 0.13s; margin-bottom: -1px;
}
.panel-tab:hover { color: #1e293b; }
.panel-tab.active { color: #334155; border-bottom-color: #1e293b; }
.panel-tab-spacer { flex: 1; }
.panel-hint { font-size: 0.65rem; color: #1e293b; font-style: italic; }

.panel-body { padding: 14px 16px; }

/* ── Palette Tab ── */
.palette-tab { display: flex; align-items: flex-start; gap: 16px; flex-wrap: wrap; }
.palette-chips { display: flex; flex-wrap: wrap; gap: 6px; flex: 1; align-content: flex-start; }
.palette-chip {
  display: flex; align-items: center; gap: 5px;
  padding: 5px 8px 5px 7px; border-radius: 8px;
  background: color-mix(in srgb, var(--c) 10%, #fff);
  border: 1px solid color-mix(in srgb, var(--c) 30%, #e2e8f0);
  cursor: grab; user-select: none;
  transition: transform 0.1s, border-color 0.1s, background 0.1s;
}
.palette-chip:hover {
  transform: translateY(-1px);
  background: color-mix(in srgb, var(--c) 18%, #fff);
  border-color: color-mix(in srgb, var(--c) 60%, #e2e8f0);
}
.palette-chip:active { cursor: grabbing; }
.p-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--c); flex-shrink: 0; }
.p-label { font-size: 0.75rem; font-weight: 500; color: color-mix(in srgb, var(--c) 70%, #1e293b); }
.p-drag { font-size: 0.65rem; color: #1e293b; margin-left: 2px; }
.p-del {
  background: none; border: none; color: #1e293b; cursor: pointer; font-size: 0.8rem;
  padding: 0 0 0 2px; line-height: 1; transition: color 0.1s;
}
.p-del:hover { color: #f43f5e; }
.p-time {
  font-family: 'DM Mono', monospace; font-size: 0.62rem;
  color: color-mix(in srgb, var(--c) 60%, #64748b); opacity: 0.85;
}
.p-time-input { font-family: 'DM Mono', monospace; font-size: 0.72rem; }

.palette-divider-v { width: 1px; background: #e2e8f0; align-self: stretch; flex-shrink: 0; }

.palette-new { display: flex; flex-direction: column; gap: 7px; min-width: 180px; }
.p-input {
  background: #fff; border: 1px solid #e2e8f0; border-radius: 7px;
  padding: 7px 10px; color: #1e293b; font-family: 'DM Sans', sans-serif; font-size: 0.75rem;
  outline: none; transition: border-color 0.13s;
}
.p-input:focus { border-color: #1e293b; }
.p-input::placeholder { color: #1e293b; }
.p-colors { display: flex; flex-wrap: wrap; gap: 5px; }
.p-swatch {
  width: 17px; height: 17px; border-radius: 50%; border: 2px solid transparent;
  cursor: pointer; transition: transform 0.1s, border-color 0.1s;
}
.p-swatch.sm { width: 14px; height: 14px; }
.p-swatch:hover { transform: scale(1.2); }
.p-swatch.selected { border-color: #1e293b; transform: scale(1.25); }
.p-add-btn {
  background: #fff; border: 1px solid #e2e8f0; color: #1e293b;
  padding: 6px 10px; border-radius: 7px;
  font-family: 'DM Sans', sans-serif; font-size: 0.72rem; font-weight: 600;
  cursor: pointer; transition: background 0.1s, border-color 0.1s;
}
.p-add-btn:hover { background: #f1f5f9; border-color: #1e293b; }
.p-add-btn:disabled { opacity: 0.3; cursor: not-allowed; }

/* ── Preacher Tab ── */
.preacher-tab { display: flex; align-items: flex-start; gap: 16px; flex-wrap: wrap; }
.preacher-form { display: flex; flex-direction: column; gap: 8px; min-width: 200px; }
.preacher-form-title { font-size: 0.72rem; font-weight: 600; color: #1e293b; text-transform: uppercase; letter-spacing: 0.08em; }
.preacher-form-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.preacher-name-input { flex: 1; min-width: 140px; }
.session-select { display: flex; gap: 3px; }
.session-btn {
  padding: 4px 9px; border-radius: 6px; border: 1px solid #e2e8f0;
  background: #fff; color: #1e293b;
  font-family: 'DM Mono', monospace; font-size: 0.68rem; font-weight: 600;
  cursor: pointer; transition: all 0.1s;
}
.session-btn:hover { border-color: #1e293b; color: #1e293b; }
.session-btn.selected { background: #eff6ff; border-color: #3b82f6; color: #2563eb; }
.preacher-swatches { display: flex; gap: 4px; }
.preacher-chip .preacher-session-badge {
  font-family: 'DM Mono', monospace; font-size: 0.55rem; font-weight: 700;
  background: rgba(0,0,0,0.1); color: rgba(0,0,0,0.5);
  padding: 1px 4px; border-radius: 3px; flex-shrink: 0;
}
.preacher-empty { font-size: 0.7rem; color: #1e293b; font-style: italic; padding: 6px 0; }

/* ── GCal Tab ── */
.gcal-tab { display: flex; gap: 20px; flex-wrap: wrap; align-items: flex-start; }
.gcal-import { display: flex; flex-direction: column; gap: 8px; min-width: 220px; flex: 1; }
.gcal-import-header { display: flex; align-items: flex-start; gap: 10px; }
.gcal-icon { font-size: 1.4rem; }
.gcal-import-title { font-size: 0.82rem; font-weight: 600; color: #334155; }
.gcal-import-sub { font-size: 0.68rem; color: #1e293b; }
.gcal-textarea {
  background: #fff; border: 1px solid #e2e8f0; border-radius: 8px;
  padding: 8px 10px; color: #1e293b; font-family: 'DM Mono', monospace; font-size: 0.65rem;
  outline: none; resize: vertical; transition: border-color 0.13s;
}
.gcal-textarea:focus { border-color: #1e293b; }
.gcal-textarea::placeholder { color: #1e293b; }
.gcal-parse-btn {
  background: #3b82f6; border: none; color: #fff;
  padding: 7px 16px; border-radius: 8px;
  font-family: 'DM Sans', sans-serif; font-size: 0.78rem; font-weight: 600;
  cursor: pointer; transition: background 0.1s; align-self: flex-start;
}
.gcal-parse-btn:hover { background: #2563eb; }
.gcal-parse-btn:disabled { opacity: 0.3; cursor: not-allowed; }
.gcal-msg { font-size: 0.72rem; }
.gcal-msg.ok { color: #16a34a; }
.gcal-msg.err { color: #dc2626; }

.gcal-divider { width: 1px; background: #e2e8f0; align-self: stretch; flex-shrink: 0; }

.gcal-manual { display: flex; flex-direction: column; gap: 8px; flex: 1; min-width: 260px; }
.gcal-manual-title { font-size: 0.72rem; font-weight: 600; color: #1e293b; text-transform: uppercase; letter-spacing: 0.08em; }
.gcal-form-row { display: flex; align-items: center; gap: 7px; flex-wrap: wrap; }
.gcal-input {
  background: #fff; border: 1px solid #e2e8f0; border-radius: 7px;
  padding: 6px 10px; color: #1e293b; font-family: 'DM Sans', sans-serif; font-size: 0.75rem;
  outline: none; transition: border-color 0.13s; flex: 1; min-width: 100px;
}
.gcal-input:focus { border-color: #1e293b; }
.gcal-input::placeholder { color: #1e293b; }
.gcal-date { flex: 0 0 130px; }
.gcal-time { flex: 0 0 100px; }
.gcal-assign-label { font-size: 0.68rem; color: #1e293b; flex-shrink: 0; }
.gcal-template-select { display: flex; flex-wrap: wrap; gap: 4px; }
.gcal-tmpl-btn {
  display: flex; align-items: center; gap: 4px;
  padding: 3px 8px; border-radius: 12px;
  background: color-mix(in srgb, var(--c) 10%, #fff);
  border: 1px solid color-mix(in srgb, var(--c) 25%, #e2e8f0);
  color: color-mix(in srgb, var(--c) 70%, #1e293b);
  font-family: 'DM Sans', sans-serif; font-size: 0.68rem; font-weight: 500;
  cursor: pointer; transition: all 0.1s;
}
.gcal-tmpl-btn:hover, .gcal-tmpl-btn.selected {
  background: color-mix(in srgb, var(--c) 20%, #fff);
  border-color: color-mix(in srgb, var(--c) 55%, #e2e8f0);
}
.custom-btn { --c: #94a3b8; }
.gtb-dot { width: 6px; height: 6px; border-radius: 50%; background: var(--c); flex-shrink: 0; }
.gcal-custom-color { display: flex; gap: 5px; align-items: center; }
.gcal-add-btn {
  background: #fff; border: 1px solid #e2e8f0; color: #1e293b;
  padding: 7px 14px; border-radius: 8px;
  font-family: 'DM Sans', sans-serif; font-size: 0.78rem; font-weight: 600;
  cursor: pointer; transition: all 0.1s; align-self: flex-start;
}
.gcal-add-btn:hover { background: #f1f5f9; border-color: #1e293b; color: #1e293b; }
.gcal-add-btn:disabled { opacity: 0.3; cursor: not-allowed; }

.gcal-list { flex: 1; min-width: 220px; display: flex; flex-direction: column; gap: 6px; }
.gcal-list-title { font-size: 0.68rem; font-weight: 600; color: #1e293b; text-transform: uppercase; letter-spacing: 0.08em; }
.gcal-list-scroll { max-height: 130px; overflow-y: auto; display: flex; flex-direction: column; gap: 3px; }
.gcal-list-item {
  display: flex; align-items: center; gap: 7px;
  padding: 5px 8px; border-radius: 7px;
  background: #fff; border: 1px solid #f1f5f9;
}
.gli-dot { width: 6px; height: 6px; border-radius: 50%; background: var(--c); flex-shrink: 0; }
.gli-title { font-size: 0.73rem; font-weight: 500; flex: 1; color: #334155; }
.gli-date { font-size: 0.65rem; color: #1e293b; flex-shrink: 0; }
.gli-time { font-family: 'DM Mono', monospace; font-size: 0.62rem; color: #1e293b; flex-shrink: 0; }
.gli-source { font-size: 0.7rem; flex-shrink: 0; }
.gli-assign { flex-shrink: 0; }
.gli-select {
  background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 5px;
  padding: 2px 6px; color: #1e293b; font-family: 'DM Sans', sans-serif; font-size: 0.65rem;
  outline: none; cursor: pointer;
}
.gli-del {
  background: none; border: none; color: #1e293b; cursor: pointer; font-size: 0.85rem;
  padding: 0; line-height: 1; transition: color 0.1s; flex-shrink: 0;
}
.gli-del:hover { color: #dc2626; }

/* ══════════════════════════════
   MODAL
══════════════════════════════ */
.modal-backdrop {
  position: fixed; inset: 0; background: rgba(15,23,42,0.45);
  backdrop-filter: blur(4px); display: flex; align-items: center; justify-content: center; z-index: 200;
}
.modal-box {
  background: #fff; border: 1px solid #e2e8f0; border-radius: 16px;
  padding: 24px; width: 340px; box-shadow: 0 20px 60px rgba(15,23,42,0.15);
}
.modal-title { font-size: 1rem; font-weight: 600; color: #1e293b; margin-bottom: 3px; }
.modal-date { font-size: 0.73rem; color: #1e293b; margin-bottom: 14px; font-style: italic; }
.modal-input {
  width: 100%; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
  padding: 9px 12px; color: #1e293b; font-family: 'DM Sans', sans-serif; font-size: 0.875rem;
  outline: none; transition: border-color 0.13s; margin-bottom: 10px;
}
.modal-input:focus { border-color: #1e293b; }
.modal-input::placeholder { color: #1e293b; }
.time-input { font-family: 'DM Mono', monospace; font-size: 0.8rem; }

.modal-templates { display: flex; align-items: center; gap: 5px; flex-wrap: wrap; margin-bottom: 10px; }
.modal-tmpl-label { font-size: 0.65rem; color: #1e293b; flex-shrink: 0; }
.modal-tmpl-btn {
  display: flex; align-items: center; gap: 3px;
  padding: 3px 8px; border-radius: 10px;
  background: color-mix(in srgb, var(--c) 10%, #fff);
  border: 1px solid color-mix(in srgb, var(--c) 25%, #e2e8f0);
  color: color-mix(in srgb, var(--c) 70%, #1e293b);
  font-family: 'DM Sans', sans-serif; font-size: 0.68rem; font-weight: 500;
  cursor: pointer; transition: all 0.1s;
}
.modal-tmpl-btn:hover, .modal-tmpl-btn.selected {
  background: color-mix(in srgb, var(--c) 20%, #fff);
  border-color: color-mix(in srgb, var(--c) 55%, #e2e8f0);
}
.mtb-dot { width: 5px; height: 5px; border-radius: 50%; background: var(--c); flex-shrink: 0; }

.color-picker { display: flex; gap: 7px; margin-bottom: 16px; flex-wrap: wrap; }
.color-swatch { width: 21px; height: 21px; border-radius: 50%; border: 2px solid transparent; cursor: pointer; transition: transform 0.1s, border-color 0.1s; }
.color-swatch:hover { transform: scale(1.15); }
.color-swatch.selected { border-color: #1e293b; transform: scale(1.2); }
.modal-actions { display: flex; gap: 8px; justify-content: flex-end; }
.btn-cancel {
  background: transparent; border: 1px solid #e2e8f0; color: #1e293b;
  padding: 7px 15px; border-radius: 8px; cursor: pointer;
  font-family: 'DM Sans', sans-serif; font-size: 0.8rem; transition: all 0.1s;
}
.btn-cancel:hover { border-color: #1e293b; color: #1e293b; }
.btn-confirm {
  background: #3b82f6; border: none; color: #fff;
  padding: 7px 18px; border-radius: 8px; cursor: pointer;
  font-family: 'DM Sans', sans-serif; font-size: 0.8rem; font-weight: 600;
  transition: background 0.1s, opacity 0.1s;
}
.btn-confirm:hover { background: #2563eb; }
.btn-confirm:disabled { opacity: 0.3; cursor: not-allowed; }
.modal-enter-active, .modal-leave-active { transition: opacity 0.18s; }
.modal-enter-from, .modal-leave-to { opacity: 0; }

/* ══════════════════════════════
   SCHOOL TERM BANDS (screen only)
══════════════════════════════ */
.term-band {
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 3px;
  border-radius: 0 0 2px 2px;
  pointer-events: none;
}
.term-school { background: rgba(16, 185, 129, 0.25); }
.term-break  { background: rgba(245, 158, 11, 0.3); }
</style>