<template>
  <div class="font-sans flex flex-col rounded-[18px] overflow-hidden border border-slate-200 bg-slate-50 text-slate-800">

    <!-- ═══════════════════════════════════
         CALENDAR
    ════════════════════════════════════ -->
    <div class="flex-1 flex flex-col p-4 pb-3 bg-slate-50">

      <!-- Print-only title -->
      <div class="hidden print:block mb-2">
        <span class="font-mono text-2xl font-medium text-slate-700 tracking-tight">Q{{ currentQuarter }} {{ currentYear }}</span>
        <span class="text-sm text-slate-500 italic ml-3">{{ months.map(m => m.name).join(' · ') }}</span>
      </div>

      <!-- Header -->
      <div class="flex items-center justify-between mb-3.5">
        <button class="action-btn" @click="prevQuarter">&#8592;</button>
        <div class="flex items-baseline gap-2.5">
          <span class="font-mono text-[1.7rem] font-medium text-slate-700 tracking-tight leading-none">Q{{ currentQuarter }}</span>
          <span class="text-base text-slate-800 font-light">{{ currentYear }}</span>
          <span class="text-[0.72rem] text-slate-800 italic">{{ months.map(m => m.name.slice(0,3)).join(' · ') }}</span>
        </div>
        <div class="flex items-center gap-2">
          <button class="action-btn" @click="nextQuarter">&#8594;</button>
          <button class="hdr-btn" @click="exportPDF" title="Export to PDF">
            <span class="text-[0.7rem]">⬇</span> PDF
          </button>
          <button class="hdr-btn border-emerald-200 text-emerald-700 hover:bg-emerald-50 hover:border-emerald-400 hover:text-emerald-800" @click="exportWord" title="Export to Word (.docx)">
            <span class="text-[0.7rem]">W</span> Word
          </button>
          <button
            class="hdr-btn border-emerald-100 text-emerald-600 hover:bg-emerald-50 hover:border-emerald-300 hover:text-emerald-800"
            :class="showSchoolTerms ? 'bg-emerald-50 !border-emerald-400 !text-emerald-800' : ''"
            @click="showSchoolTerms = !showSchoolTerms"
            title="Toggle school term highlights"
          >🏫 Terms</button>
          <button class="hdr-btn border-yellow-200 text-yellow-900 hover:bg-yellow-50 hover:border-yellow-400 hover:text-yellow-900" @click="loadDefaultEvents" title="Load default recurring events">
            ✦ Defaults
          </button>
        </div>
      </div>

      <!-- Calendar Grid -->
      <div class="overflow-x-auto">
        <div
          class="cal-grid grid gap-px bg-slate-200 border border-slate-200 rounded-xl overflow-hidden w-full"
          :style="{ gridTemplateColumns: `56px repeat(${allWeeks.length}, minmax(80px, 1fr))` }"
        >
          <!-- Row 0: Month spans -->
          <div class="bg-slate-100 p-1.5 text-[0.48rem] text-slate-800 uppercase tracking-wide flex items-end leading-relaxed">Wk →<br/>Day ↓</div>
          <template v-for="(month, mi) in monthsWithCounts" :key="'mh-' + mi">
            <div
              class="bg-slate-100 px-2 py-1.5 text-[0.65rem] font-semibold tracking-widest uppercase text-center text-slate-800 border-b border-slate-200"
              :style="{ gridColumn: `span ${month.weekCount}` }"
            >{{ month.name }}</div>
          </template>

          <!-- Row 1: Week headers -->
          <div class="bg-slate-100"></div>
          <div
            v-for="(week, wi) in allWeeks" :key="'wh-' + wi"
            class="bg-slate-100 px-1 py-1.5 text-center"
            :class="{ 'bg-blue-50': isCurrentWeek(week), 'border-l-2 border-l-slate-200': week.isMonthBoundary }"
          >
            <span class="block font-mono text-[0.58rem] text-slate-800 leading-none mb-0.5">{{ week.weekNumber }}</span>
            <span class="block text-[0.54rem] text-slate-800 whitespace-nowrap">{{ formatWeekStart(week.startDate) }}</span>
          </div>

          <!-- Day rows -->
          <template v-for="(day, di) in dayNames" :key="'row-' + di">
            <div class="bg-slate-100 px-1 flex items-center justify-center min-h-[54px]">
              <span class="font-mono text-[0.6rem] text-slate-800 font-medium">{{ day.short }}</span>
            </div>
            <div
              v-for="(week, wi) in allWeeks" :key="'cell-' + di + '-' + wi"
              class="day-cell bg-white min-h-[54px] min-h-[75px] p-0.5 relative cursor-pointer transition-colors duration-75 overflow-hidden hover:bg-slate-50"
              :class="{
                'bg-blue-50 today-cell': isToday(week.startDate, di),
                'opacity-25 pointer-events-none': !isInQuarter(week.startDate, di),
                'drag-over-cell': dragOverKey === makeCellKey(week.startDate, di),
                'border-l-2 border-l-slate-200': week.isMonthBoundary
              }"
              @click="openAddEvent(week.startDate, di)"
              @dragover.prevent="dragOverKey = makeCellKey(week.startDate, di)"
              @dragleave="dragOverKey = null"
              @drop.prevent="onDropToCell($event, week.startDate, di)"
            >
              <span
                class="font-mono text-[0.56rem] block mb-0.5 leading-none"
                :class="isToday(week.startDate, di) ? 'text-blue-500 font-semibold' : 'text-slate-800'"
              >{{ getCellDate(week.startDate, di) }}</span>

              <div
                v-if="getSchoolPeriodForCell(week.startDate, di)"
                class="term-band absolute bottom-0 left-0 right-0 h-[3px] rounded-b-sm pointer-events-none"
                :class="getSchoolPeriodForCell(week.startDate, di).type === 'break' ? 'term-break' : 'term-school'"
                :title="getSchoolPeriodForCell(week.startDate, di).label"
              ></div>

              <div class="flex flex-col gap-px">
                <div
                  v-for="ev in getEventsForCell(week.startDate, di)" :key="ev.id"
                  class="event-chip flex items-center gap-0.5 px-1 py-0.5 rounded-[3px] cursor-grab transition-all max-w-full hover:opacity-75 hover:scale-[1.03] active:cursor-grabbing"
                  :style="{ '--ec': ev.color }"
                  :class="{ 'border-dashed': ev.source === 'gcal' }"
                  :title="ev.title + (ev.time ? ' @ ' + ev.time : '') + '\nClick to remove'"
                  draggable="true"
                  @dragstart="onEventDragStart($event, ev)"
                  @click.stop="removeEvent(ev)"
                >
                  <span v-if="ev.time" class="text-[0.5rem] text-black/95 bg-black/[0.07] rounded px-0.5 shrink-0 font-semibold tracking-tight">{{ ev.time }}</span>
                  <span class="text-[0.56rem] font-semibold text-slate-800 truncate pointer-events-none flex-1">{{ ev.title }}</span>
                  <span v-if="ev.source === 'gcal'" class="w-1 h-1 rounded-full bg-black/25 shrink-0" title="From Google Calendar"></span>
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
    <div class="border-t border-slate-200 bg-slate-50 shrink-0">

      <!-- Tab bar -->
      <div class="flex items-center border-b border-slate-200 px-4">
        <button
          v-for="tab in ['palette', 'preacher', 'gcal']" :key="tab"
          class="bg-transparent border-none border-b-2 border-b-transparent -mb-px px-4 py-2.5 cursor-pointer font-sans text-[0.78rem] font-medium transition-colors duration-150 text-slate-800"
          :class="activeTab === tab ? 'text-slate-700 !border-b-slate-800' : 'hover:text-slate-700'"
          @click="activeTab = tab"
        >
          <span v-if="tab === 'palette'">🎨 Event Types</span>
          <span v-if="tab === 'preacher'">🎤 Guest Preachers</span>
          <span v-if="tab === 'gcal'">📅 Google Calendar</span>
        </button>
        <div class="flex-1"></div>
        <span class="text-[0.65rem] text-slate-800 italic">Drag events onto the calendar · Click an event to remove it</span>
      </div>

      <!-- ── TAB: Palette ── -->
      <div v-if="activeTab === 'palette'" class="p-3.5 flex items-start gap-4 flex-wrap min-h-[200px]">
        <div class="flex flex-wrap gap-1.5 flex-1 content-start">
          <div
            v-for="tmpl in eventTemplates" :key="tmpl.id"
            class="palette-chip flex items-center gap-1.5 py-1.5 px-2 rounded-lg cursor-grab select-none transition-all hover:-translate-y-px active:cursor-grabbing"
            :style="{ '--c': tmpl.color }"
            draggable="true"
            @dragstart="onPaletteDragStart($event, tmpl)"
          >
            <span class="p-dot w-[7px] h-[7px] rounded-full shrink-0" :style="{ background: tmpl.color }"></span>
            <span class="p-label text-[0.75rem] font-medium">{{ tmpl.title }}</span>
            <span v-if="tmpl.time" class="font-mono text-[0.62rem] opacity-95" :style="{ color: `color-mix(in srgb, ${tmpl.color} 60%, #64748b)` }">{{ tmpl.time }}</span>
            <span class="text-[0.65rem] text-slate-800 ml-0.5">⠿</span>
            <button class="bg-transparent border-none text-slate-800 cursor-pointer text-[0.8rem] pl-0.5 leading-none transition-colors hover:text-rose-500" @click.stop="deleteTemplate(tmpl)" title="Remove type">×</button>
          </div>
        </div>

        <div class="w-px bg-slate-200 self-stretch shrink-0"></div>

        <div class="flex flex-col gap-1.5 min-w-[180px]">
          <input v-model="newTmpl.title" class="field-input" placeholder="New type name..." @keyup.enter="addTemplate" />
          <TimePicker v-model="newTmpl.time" />
          <div class="flex flex-wrap gap-1.5">
            <button
              v-for="c in eventColors" :key="c"
              class="w-[17px] h-[17px] rounded-full border-2 border-transparent cursor-pointer transition-all hover:scale-125"
              :style="{ background: c }"
              :class="newTmpl.color === c ? 'border-slate-800 scale-[1.25]' : ''"
              @click="newTmpl.color = c"
            ></button>
          </div>
          <button class="field-btn" @click="addTemplate" :disabled="!newTmpl.title.trim()">+ Add Type</button>
        </div>
      </div>

      <!-- ── TAB: Guest Preachers ── -->
      <div v-if="activeTab === 'preacher'" class="p-3.5 flex items-start gap-4 flex-wrap">
        <div class="flex flex-col gap-2 min-w-[200px]">
          <div class="text-[0.72rem] font-semibold text-slate-800 uppercase tracking-widest">Add Guest Preacher</div>
          <div class="flex items-center gap-2 flex-wrap">
            <input v-model="preacherNew.name" class="field-input flex-1 min-w-[140px]" placeholder="Preacher name..." @keyup.enter="addPreacher" />
            <div class="flex gap-0.5">
              <button
                v-for="s in ['AM', 'PM', 'All Day']" :key="s"
                class="px-2.5 py-1 rounded-md border font-mono text-[0.68rem] font-semibold cursor-pointer transition-all"
                :class="preacherNew.session === s
                  ? 'bg-blue-50 border-blue-400 text-blue-700'
                  : 'bg-white border-slate-200 text-slate-800 hover:border-slate-800'"
                @click="preacherNew.session = s"
              >{{ s }}</button>
            </div>
            <div class="flex gap-1">
              <button
                v-for="c in preacherColors" :key="c"
                class="w-3.5 h-3.5 rounded-full border-2 border-transparent cursor-pointer transition-all hover:scale-125"
                :style="{ background: c }"
                :class="preacherNew.color === c ? 'border-slate-800 scale-[1.25]' : ''"
                @click="preacherNew.color = c"
              ></button>
            </div>
            <button class="field-btn" @click="addPreacher" :disabled="!preacherNew.name.trim()">+ Add</button>
          </div>
        </div>

        <div class="w-px bg-slate-200 self-stretch shrink-0"></div>

        <div class="flex flex-wrap gap-1.5 flex-1">
          <div
            v-for="p in sortedPreachers" :key="p.id"
            class="palette-chip flex items-center gap-1.5 py-1.5 px-2 rounded-lg cursor-grab select-none transition-all hover:-translate-y-px active:cursor-grabbing"
            :style="{ '--c': p.color }"
            draggable="true"
            @dragstart="onPaletteDragStart($event, { ...p, title: p.name + (p.session !== 'All Day' ? ' (' + p.session + ')' : '') })"
          >
            <span class="w-[7px] h-[7px] rounded-full shrink-0" :style="{ background: p.color }"></span>
            <span class="p-label text-[0.75rem] font-medium">{{ p.name }}</span>
            <span class="font-mono text-[0.55rem] font-bold bg-black/10 text-black/50 px-1 py-px rounded shrink-0">{{ p.session }}</span>
            <span class="text-[0.65rem] text-slate-800 ml-0.5">⠿</span>
            <button class="bg-transparent border-none text-slate-800 cursor-pointer text-[0.8rem] pl-0.5 leading-none transition-colors hover:text-rose-500" @click.stop="removePreacher(p)" title="Remove">×</button>
          </div>
          <div v-if="!preachers.length" class="text-[0.7rem] text-slate-800 italic py-1.5">
            Add a guest preacher above, then drag them onto the calendar
          </div>
        </div>
      </div>

      <!-- ── TAB: Google Calendar ── -->
      <div v-if="activeTab === 'gcal'" class="p-3.5 flex gap-5 flex-wrap items-start">

        <div class="flex flex-col gap-2 min-w-[220px] flex-1">
          <div class="flex items-start gap-2.5">
            <span class="text-[1.4rem]">📅</span>
            <div>
              <div class="text-[0.82rem] font-semibold text-slate-700">Import from Google Calendar</div>
              <div class="text-[0.68rem] text-slate-800">Paste ICS/CSV export, or add events manually below</div>
            </div>
          </div>
          <textarea
            v-model="gcalPaste"
            class="bg-white border border-slate-200 rounded-lg px-2.5 py-2 text-slate-800 font-mono text-[0.65rem] outline-none resize-y transition-colors focus:border-slate-800 placeholder-slate-800"
            placeholder="Paste Google Calendar .ics content here..."
            rows="3"
          ></textarea>
          <button class="self-start px-4 py-1.5 rounded-lg bg-blue-500 text-white font-sans text-[0.78rem] font-semibold cursor-pointer border-none transition-colors hover:bg-blue-600 disabled:opacity-30 disabled:cursor-not-allowed" @click="parseICS" :disabled="!gcalPaste.trim()">
            Parse &amp; Import Events
          </button>
          <span v-if="gcalImportMsg" class="text-[0.72rem]" :class="gcalImportOk ? 'text-emerald-600' : 'text-red-600'">{{ gcalImportMsg }}</span>
        </div>

        <div class="w-px bg-slate-200 self-stretch shrink-0"></div>

        <div class="flex flex-col gap-2 flex-1 min-w-[260px]">
          <div class="text-[0.72rem] font-semibold text-slate-800 uppercase tracking-widest">Add event manually</div>
          <div class="flex items-center gap-1.5 flex-wrap">
            <input v-model="gcalNew.title" class="field-input flex-1 min-w-[100px]" placeholder="Event title" />
            <input v-model="gcalNew.date"  class="field-input w-[130px]" type="date" />
            <input v-model="gcalNew.time"  class="field-input w-[100px]" type="time" placeholder="Time (opt.)" />
          </div>
          <div class="flex items-center gap-1.5 flex-wrap">
            <span class="text-[0.68rem] text-slate-800 shrink-0">Assign type:</span>
            <div class="flex flex-wrap gap-1">
              <button
                v-for="tmpl in eventTemplates" :key="tmpl.id"
                class="gcal-tmpl-btn flex items-center gap-1 px-2 py-0.5 rounded-xl font-sans text-[0.68rem] font-medium cursor-pointer transition-all"
                :style="{ '--c': tmpl.color }"
                :class="{ 'gcal-tmpl-selected': gcalNew.templateId === tmpl.id }"
                @click="gcalNew.templateId = (gcalNew.templateId === tmpl.id ? null : tmpl.id)"
              >
                <span class="w-1.5 h-1.5 rounded-full shrink-0" :style="{ background: tmpl.color }"></span>{{ tmpl.title }}
              </button>
              <button
                class="gcal-tmpl-btn flex items-center gap-1 px-2 py-0.5 rounded-xl font-sans text-[0.68rem] font-medium cursor-pointer transition-all"
                style="--c: #94a3b8"
                :class="{ 'gcal-tmpl-selected': gcalNew.templateId === 'custom' }"
                @click="gcalNew.templateId = 'custom'"
              >Custom</button>
            </div>
            <div v-if="gcalNew.templateId === 'custom'" class="flex gap-1.5 items-center">
              <button
                v-for="c in eventColors" :key="c"
                class="w-3.5 h-3.5 rounded-full border-2 border-transparent cursor-pointer transition-all hover:scale-125"
                :style="{ background: c }"
                :class="gcalNew.customColor === c ? 'border-slate-800 scale-[1.25]' : ''"
                @click="gcalNew.customColor = c"
              ></button>
            </div>
          </div>
          <button class="field-btn self-start" @click="addGCalEvent" :disabled="!gcalNew.title.trim() || !gcalNew.date">
            + Add to Calendar
          </button>
        </div>

        <div v-if="gcalEvents.length" class="flex-1 min-w-[220px] flex flex-col gap-1.5">
          <div class="text-[0.68rem] font-semibold text-slate-800 uppercase tracking-widest">Imported events ({{ gcalEvents.length }})</div>
          <div class="max-h-[130px] overflow-y-auto flex flex-col gap-0.5">
            <div v-for="ev in gcalEvents" :key="ev.id" class="flex items-center gap-1.5 px-2 py-1.5 rounded-lg bg-white border border-slate-100" :style="{ '--c': ev.color }">
              <span class="w-1.5 h-1.5 rounded-full shrink-0" :style="{ background: ev.color }"></span>
              <span class="text-[0.73rem] font-medium flex-1 text-slate-700">{{ ev.title }}</span>
              <span class="text-[0.65rem] text-slate-800 shrink-0">{{ formatGCalDate(ev.date) }}</span>
              <span v-if="ev.time" class="font-mono text-[0.62rem] text-slate-800 shrink-0">{{ ev.time }}</span>
              <span class="text-[0.7rem] shrink-0">{{ ev.source === 'gcal' ? '📅' : '✏️' }}</span>
              <div class="shrink-0">
                <select class="bg-slate-50 border border-slate-200 rounded px-1.5 py-0.5 text-slate-800 font-sans text-[0.65rem] outline-none cursor-pointer" :value="ev.assignedTemplateId" @change="assignTemplate(ev, $event.target.value)">
                  <option value="">— type —</option>
                  <option v-for="tmpl in eventTemplates" :key="tmpl.id" :value="tmpl.id">{{ tmpl.title }}</option>
                </select>
              </div>
              <button class="bg-transparent border-none text-slate-800 cursor-pointer text-[0.85rem] leading-none transition-colors hover:text-red-600" @click="removeGCalEvent(ev)">×</button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- ═══════════════════════════════════
         ADD EVENT MODAL
    ════════════════════════════════════ -->
    <transition name="modal">
      <div v-if="modal.open" class="fixed inset-0 bg-slate-900/45 backdrop-blur-sm flex items-center justify-center z-[200]" @click.self="modal.open = false">
        <div class="bg-white border border-slate-200 rounded-2xl p-6 w-[340px] shadow-[0_20px_60px_rgba(15,23,42,0.15)]">
          <h3 class="text-base font-semibold text-slate-800 mb-0.5">Add Event</h3>
          <p class="text-[0.73rem] text-slate-800 italic mb-3.5">{{ formatModalDate(modal.date) }}</p>
          <input v-model="modal.title" class="field-input w-full mb-2.5" placeholder="Event title..." @keyup.enter="confirmAddEvent" ref="modalInput" />
          <TimePicker v-model="modal.time" className="!mb-4" />
          <div class="flex items-center gap-1.5 flex-wrap mb-2.5">
            <span class="text-[0.65rem] text-slate-800 shrink-0">Quick type:</span>
            <button
              v-for="tmpl in eventTemplates" :key="tmpl.id"
              class="gcal-tmpl-btn flex items-center gap-1 px-2 py-0.5 rounded-[10px] font-sans text-[0.68rem] font-medium cursor-pointer transition-all"
              :style="{ '--c': tmpl.color }"
              :class="{ 'gcal-tmpl-selected': modal.templateId === tmpl.id }"
              @click="applyTemplate(tmpl)"
            >
              <span class="w-[5px] h-[5px] rounded-full shrink-0" :style="{ background: tmpl.color }"></span>{{ tmpl.title }}
            </button>
          </div>
          <div class="flex gap-1.5 mb-4 flex-wrap">
            <button
              v-for="c in eventColors" :key="c"
              class="w-[21px] h-[21px] rounded-full border-2 border-transparent cursor-pointer transition-all hover:scale-[1.15]"
              :style="{ background: c }"
              :class="modal.color === c ? 'border-slate-800 scale-[1.2]' : ''"
              @click="modal.color = c"
            ></button>
          </div>
          <div class="flex gap-2 justify-end">
            <button class="bg-transparent border border-slate-200 text-slate-800 px-4 py-1.5 rounded-lg cursor-pointer font-sans text-[0.8rem] transition-all hover:border-slate-800" @click="modal.open = false">Cancel</button>
            <button class="bg-blue-500 border-none text-white px-4 py-1.5 rounded-lg cursor-pointer font-sans text-[0.8rem] font-semibold transition-all hover:bg-blue-600 disabled:opacity-30 disabled:cursor-not-allowed" @click="confirmAddEvent" :disabled="!modal.title.trim()">Add Event</button>
          </div>
        </div>
      </div>
    </transition>

  </div>
</template>

<script>
import { zipSync, strToU8 } from 'fflate'
import schoolTermsData from '~/data/schoolTerms.js'
import TimePicker from './time-component.vue'

let _uid = 1

export default {
  name: 'QuarterlyCalendar',
  components: { TimePicker },
  data() {
    const today = new Date()
    return {
      today: new Date(today.getFullYear(), today.getMonth(), today.getDate()),
      currentYear: today.getFullYear(),
      currentQuarter: Math.ceil((today.getMonth() + 1) / 3),
      activeTab: 'palette',
      events: [],
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
        { id: _uid++, title: 'Prayer Meeting',       time: '19:30', color: '#6366f1' },
        { id: _uid++, title: 'Discoverers',          time: '19:30', color: '#f43f5e' },
        { id: _uid++, title: 'First Steps',          time: '9:00',  color: '#ffeffe' },
        { id: _uid++, title: 'Coffee Morning',       time: '10:00', color: '#0a2000' },
        { id: _uid++, title: 'Small Groups',         time: '10:00', color: '#10b981' },
        { id: _uid++, title: 'Small Groups',         time: '19:30', color: '#10b981' },
        { id: _uid++, title: 'Small Groups',         time: '19:45', color: '#10b981' },
        { id: _uid++, title: 'Members Meeting',      time: '',      color: '#14b8a6' },
        { id: _uid++, title: "Men's Meeting",        time: '',      color: '#1401f0' },
        { id: _uid++, title: 'Ladies Meeting',       time: '',      color: '#ec4899' },
        { id: _uid++, title: '(Rooted and Grounded)',time: '',      color: '#14b8a6' },
        { id: _uid++, title: 'Joint Meeting',        time: '',      color: '#14b8a6' },
        { id: _uid++, title: 'Pastor Paul',          time: '',      color: '#60a5fa' },
      ],
      newTmpl: { title: '', time: '', color: '#6366f1' },
      schoolPeriods: schoolTermsData,
      showSchoolTerms: true,
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
      cur.setDate(cur.getDate() - dow)
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
    prevQuarter() { this.currentQuarter--; if (this.currentQuarter < 1) { this.currentQuarter = 4; this.currentYear-- } },
    nextQuarter() { this.currentQuarter++; if (this.currentQuarter > 4) { this.currentQuarter = 1; this.currentYear++ } },

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
    onPaletteDragStart(e, tmpl) {
      this.draggingTemplate = tmpl; this.draggingEvent = null; this.draggingGCal = null
      e.dataTransfer.effectAllowed = 'copy'
    },
    onEventDragStart(e, ev) {
      this.draggingEvent = ev; this.draggingTemplate = null; this.draggingGCal = null
      e.dataTransfer.effectAllowed = 'move'; e.stopPropagation()
    },
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
      this.events.push({ id: _uid++, key: this.dateToKey(d), title: this.modal.title.trim(), time: this.modal.time || '', color: this.modal.color, source: 'manual', date: new Date(d) })
      this.modal.open = false
    },
    removeEvent(ev) {
      const idx = this.events.indexOf(ev); if (idx > -1) this.events.splice(idx, 1)
    },
    addPreacher() {
      if (!this.preacherNew.name.trim()) return
      this.preachers.push({ id: _uid++, name: this.preacherNew.name.trim(), session: this.preacherNew.session, color: this.preacherNew.color, time: '' })
      this.preacherNew = { name: '', session: 'AM', color: this.preacherColors[0] }
    },
    removePreacher(p) {
      const i = this.preachers.indexOf(p); if (i > -1) this.preachers.splice(i, 1)
    },
    addTemplate() {
      if (!this.newTmpl.title.trim()) return
      this.eventTemplates.push({ id: _uid++, title: this.newTmpl.title.trim(), time: this.newTmpl.time || '', color: this.newTmpl.color })
      this.newTmpl = { title: '', time: '', color: this.eventColors[0] }
    },
    deleteTemplate(tmpl) {
      const idx = this.eventTemplates.indexOf(tmpl); if (idx > -1) this.eventTemplates.splice(idx, 1)
    },
    addGCalEvent() {
      if (!this.gcalNew.title.trim() || !this.gcalNew.date) return
      const tmpl = this.gcalNew.templateId && this.gcalNew.templateId !== 'custom'
        ? this.eventTemplates.find(t => t.id === this.gcalNew.templateId) : null
      const color = tmpl ? tmpl.color : (this.gcalNew.customColor || '#888')
      const ev = { id: _uid++, title: this.gcalNew.title.trim(), date: this.gcalNew.date, time: this.gcalNew.time || '', color, source: 'manual-gcal', assignedTemplateId: tmpl ? tmpl.id : null }
      this.gcalEvents.push(ev)
      this.placeGCalEventOnGrid(ev)
      this.gcalNew = { title: '', date: '', time: '', templateId: null, customColor: '#6366f1' }
    },
    placeGCalEventOnGrid(ev) {
      if (!ev.date) return
      const d = new Date(ev.date + 'T00:00:00')
      const key = this.dateToKey(d)
      const exists = this.events.find(e => e.gcalId === ev.id)
      if (exists) { exists.key = key; exists.time = ev.time; exists.color = ev.color; return }
      this.events.push({ id: _uid++, gcalId: ev.id, key, title: ev.title, time: ev.time || '', color: ev.color, source: 'gcal', date: d })
    },
    assignTemplate(ev, templateId) {
      ev.assignedTemplateId = templateId || null
      const tmpl = templateId ? this.eventTemplates.find(t => String(t.id) === String(templateId)) : null
      if (tmpl) {
        ev.color = tmpl.color
        const placed = this.events.find(e => e.gcalId === ev.id)
        if (placed) placed.color = tmpl.color
      }
    },
    removeGCalEvent(ev) {
      const idx = this.gcalEvents.indexOf(ev); if (idx > -1) this.gcalEvents.splice(idx, 1)
      const pi = this.events.findIndex(e => e.gcalId === ev.id)
      if (pi > -1) this.events.splice(pi, 1)
    },
    parseICS() {
      this.gcalImportMsg = ''; this.gcalImportOk = false
      const text = this.gcalPaste
      const events = []
      const veventBlocks = text.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || []
      if (veventBlocks.length === 0) { this.gcalImportMsg = 'No events found. Make sure you pasted valid .ics content.'; return }
      veventBlocks.forEach(block => {
        const get = (key) => { const m = block.match(new RegExp(key + '[^:]*:(.+)')); return m ? m[1].replace(/\r/g, '').trim() : '' }
        const summary = get('SUMMARY'), dtstart = get('DTSTART')
        if (!summary || !dtstart) return
        const dateStr = dtstart.replace(/T.*/, '').replace(/[^0-9]/g,'')
        if (dateStr.length < 8) return
        const isoDate = `${dateStr.slice(0,4)}-${dateStr.slice(4,6)}-${dateStr.slice(6,8)}`
        let time = ''
        const timePart = dtstart.match(/T(\d{2})(\d{2})/)
        if (timePart) time = `${timePart[1]}:${timePart[2]}`
        events.push({ id: _uid++, title: summary, date: isoDate, time, color: this.eventColors[0], source: 'gcal', assignedTemplateId: null })
      })
      if (events.length === 0) { this.gcalImportMsg = 'Could not parse any events. Check the format.'; return }
      events.forEach(ev => { this.gcalEvents.push(ev); this.placeGCalEventOnGrid(ev) })
      this.gcalPaste = ''
      this.gcalImportMsg = `✓ Imported ${events.length} event${events.length > 1 ? 's' : ''}`
      this.gcalImportOk = true
    },

    exportWord() {
      const weeks = this.allWeeks
      const months = this.monthsWithCounts
      const quarterTitle = `Q${this.currentQuarter} ${this.currentYear} — ${this.months.map(m => m.name).join(', ')}`
      const printDate = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })
      const esc = (s) => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;')
      const hex = (color) => color.replace('#','').toUpperCase()
      const MARGIN = 454, CONTENT_W = 16838 - MARGIN * 2, DAY_COL = 560
      const weekColW = Math.floor((CONTENT_W - DAY_COL) / weeks.length)
      const lastColW = CONTENT_W - DAY_COL - weekColW * (weeks.length - 1)
      const colW = (wi) => wi === weeks.length - 1 ? lastColW : weekColW
      const borderXml = (leftThick) => `<w:tcBorders><w:top w:val="single" w:sz="4" w:color="AAAAAA"/><w:bottom w:val="single" w:sz="4" w:color="AAAAAA"/><w:left w:val="single" w:sz="${leftThick?12:4}" w:color="${leftThick?'666666':'AAAAAA'}"/><w:right w:val="single" w:sz="4" w:color="AAAAAA"/></w:tcBorders>`
      const cellPr = (width, fill, leftThick, vAlign, span) => `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${span>1?`<w:gridSpan w:val="${span}"/>`:''}${fill?`<w:shd w:val="clear" w:fill="${fill}"/>`:''}${borderXml(leftThick)}<w:tcMar><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>${vAlign?`<w:vAlign w:val="${vAlign}"/>`:''}</w:tcPr>`
      const para = (text, opts = {}) => {
        const jc = opts.align ? `<w:jc w:val="${opts.align}"/>` : ''
        const ind = opts.indent ? `<w:ind w:left="${opts.indent}"/>` : ''
        const pBdr = opts.leftColor ? `<w:pBdr><w:left w:val="single" w:sz="8" w:space="2" w:color="${hex(opts.leftColor)}"/></w:pBdr>` : ''
        const sz = opts.size || 18
        const runs = (opts.runs || [{ text, ...opts }]).map(r => `<w:r><w:rPr><w:sz w:val="${r.size||sz}"/><w:szCs w:val="${r.size||sz}"/>${r.bold?'<w:b/>':''}${r.color?`<w:color w:val="${r.color}"/>`:''}${r.italic?'<w:i/>':''}</w:rPr><w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`).join('')
        return `<w:p><w:pPr>${jc}${ind}${pBdr}</w:pPr>${runs}</w:p>`
      }
      const cell = (content, width, fill, leftThick = false, vAlign = '', span = 1) => `<w:tc>${cellPr(width, fill, leftThick, vAlign, span)}${content||'<w:p/>'}</w:tc>`
      const monthFills = ['EDE9FE','DBEAFE','D1FAE5'], monthColors = ['5B21B6','1D4ED8','065F46']
      let monthHeaderRow = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('',DAY_COL,'F0F0F0')}`
      months.forEach((m, mi) => {
        const spanCount = m.weekCount
        const width = weekColW * spanCount + (mi === months.length - 1 ? lastColW - weekColW : 0)
        monthHeaderRow += cell(para(m.name.toUpperCase(),{align:'center',bold:true,size:16,color:monthColors[mi]}),width,monthFills[mi],mi>0,'',spanCount)
      })
      monthHeaderRow += '</w:tr>'
      let weekHeaderRow = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('',DAY_COL,'F0F0F0')}`
      weeks.forEach((w, wi) => {
        const label = w.startDate.toLocaleDateString('en-GB',{day:'numeric',month:'short'})
        weekHeaderRow += cell(para('',{align:'center',runs:[{text:`Wk ${w.weekNumber}`,size:14,bold:true,color:'444444'}]})+para('',{align:'center',runs:[{text:label,size:12,color:'888888'}]}),colW(wi),'F8F8F8',w.isMonthBoundary)
      })
      weekHeaderRow += '</w:tr>'
      const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
      const dayRows = DAY_NAMES.map((dayName, di) => {
        let row = `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="0" w:hRule="auto"/></w:trPr>`
        row += cell(para(dayName,{align:'center',bold:true,size:16,color:'333333'}),DAY_COL,'F8F8F8',false,'center')
        weeks.forEach((week, wi) => {
          const d = new Date(week.startDate); d.setDate(d.getDate() + di)
          const inQ = this.months.some(mo => mo.index === d.getMonth())
          const key = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`
          const evs = inQ ? this.events.filter(e => e.key === key).sort((a,b) => (a.time||'').localeCompare(b.time||'')) : []
          let cellContent = ''
          if (inQ) {
            cellContent += para(String(d.getDate()),{size:12,color:'BBBBBB'})
            evs.forEach(ev => { cellContent += para('',{leftColor:ev.color,indent:80,runs:ev.time?[{text:ev.time+' ',size:14,color:'666666'},{text:ev.title,size:14,bold:true,color:'111111'}]:[{text:ev.title,size:14,color:'222222'}]}) })
          }
          row += cell(cellContent||'<w:p/>',colW(wi),inQ?'FFFFFF':'F5F5F5',week.isMonthBoundary)
        })
        return row + '</w:tr>'
      }).join('')
      const gridCols = [DAY_COL,...weeks.map((_,wi) => colW(wi))].map(w => `<w:gridCol w:w="${w}"/>`).join('')
      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:spacing w:after="80"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>${esc(quarterTitle)}</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="888888"/></w:rPr><w:t>Printed ${esc(printDate)}</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="${CONTENT_W}" w:type="dxa"/><w:tblBorders><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBorders><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${gridCols}</w:tblGrid>${monthHeaderRow}${weekHeaderRow}${dayRows}</w:tbl><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="0" w:footer="0"/></w:sectPr></w:body></w:document>`
      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
      const wordRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`
      const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/></w:settings>`
      const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`
      const zipped = zipSync({ '[Content_Types].xml': strToU8(contentTypesXml), '_rels/.rels': strToU8(relsXml), 'word/document.xml': strToU8(docXml), 'word/_rels/document.xml.rels': strToU8(wordRelsXml), 'word/settings.xml': strToU8(settingsXml) }, { level: 6 })
      const blob = new Blob([zipped], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a'); a.href = url; a.download = `Q${this.currentQuarter}-${this.currentYear}-calendar.docx`; a.click(); URL.revokeObjectURL(url)
    },

    exportPDF() {
      const weeks = this.allWeeks
      const dayNames = this.dayNames
      const monthsInfo = this.monthsWithCounts
      const getEvents = (weekStart, di) => {
        const d = new Date(weekStart); d.setDate(d.getDate() + di)
        const key = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`
        return this.events.filter(e => e.key === key).sort((a,b) => (a.time||'').localeCompare(b.time||''))
      }
      const getCellDate = (weekStart, di) => { const d = new Date(weekStart); d.setDate(d.getDate() + di); return d.getDate() }
      const isInQuarter = (weekStart, di) => { const d = new Date(weekStart); d.setDate(d.getDate() + di); return this.months.some(mo => mo.index === d.getMonth()) }
      let monthHeaders = '<th style="width:36px;border:1px solid #999;padding:3px 4px;font-size:7pt;text-align:center;background:#f0f0f0;"></th>'
      monthsInfo.forEach(m => { monthHeaders += `<th colspan="${m.weekCount}" style="border:1px solid #999;padding:4px 6px;font-size:8pt;font-weight:bold;text-align:center;background:#e8e8e8;letter-spacing:0.05em;">${m.name.toUpperCase()}</th>` })
      let weekHeaders = '<th style="border:1px solid #999;padding:3px 2px;font-size:6pt;text-align:center;background:#f0f0f0;"></th>'
      weeks.forEach(w => {
        const label = w.startDate.toLocaleDateString('en-GB',{day:'numeric',month:'short'})
        const boundary = w.isMonthBoundary ? 'border-left:2px solid #555;' : ''
        weekHeaders += `<th style="border:1px solid #bbb;${boundary}padding:3px 2px;font-size:6pt;font-weight:600;text-align:center;background:#f8f8f8;min-width:0;"><div style="font-size:6.5pt;color:#555;">Wk ${w.weekNumber}</div><div style="font-size:5.5pt;color:#888;">${label}</div></th>`
      })
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
          cells += `<td style="border:1px solid #ddd;${boundary}${cellBg}padding:2px 3px;vertical-align:top;min-width:0;">${inQ ? `<div style="font-size:5.5pt;color:#aaa;line-height:1;margin-bottom:1px;">${dateNum}</div>${evHTML}` : ''}</td>`
        })
        dayRows += `<tr>${cells}</tr>`
      })
      const quarterTitle = `Q${this.currentQuarter} ${this.currentYear} — ${this.months.map(m => m.name).join(', ')}`
      const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${quarterTitle}</title><style>@page{size:A4 landscape;margin:8mm}*{box-sizing:border-box;margin:0;padding:0}html,body{height:100%;font-family:Arial,Helvetica,sans-serif;font-size:8pt;color:#000;background:#fff}.page{display:flex;flex-direction:column;height:100%}h1{font-size:11pt;font-weight:bold;margin-bottom:2px;color:#111;flex-shrink:0}.sub{font-size:7pt;color:#666;margin-bottom:6px;flex-shrink:0}.table-wrap{flex:1;display:flex;flex-direction:column;overflow:hidden}table{width:100%;height:100%;border-collapse:collapse;table-layout:fixed}tbody{height:100%}tbody tr{height:calc(100% / 7)}td,th{overflow:hidden;word-break:break-word;vertical-align:top}@media print{html,body{height:100%;margin:0}}</style></head><body><div class="page"><h1>${quarterTitle}</h1><div class="sub">Printed ${new Date().toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'})}</div><div class="table-wrap"><table><thead><tr>${monthHeaders}</tr><tr>${weekHeaders}</tr></thead><tbody>${dayRows}</tbody></table></div></div></body></html>`
      const win = window.open('','_blank','width=1200,height=800')
      if (!win) { alert('Please allow pop-ups to export PDF'); return }
      win.document.write(html); win.document.close(); win.focus()
      setTimeout(() => win.print(), 400)
    },

    loadDefaultEvents() {
      const year = this.currentYear
      const qStart = this.quarterStartMonth
      const add = (date, title, time, color) => {
        const key = this.dateToKey(date)
        if (this.events.find(e => e.key === key && e.title === title && e.time === time)) return
        this.events.push({ id: _uid++, key, title, time, color, source: 'manual', date: new Date(date) })
      }
      for (let m = 0; m < 3; m++) {
        const monthIdx = qStart + m
        const daysInMonth = new Date(year, monthIdx + 1, 0).getDate()
        let lastWed = new Date(year, monthIdx, daysInMonth)
        while (lastWed.getDay() !== 3) lastWed.setDate(lastWed.getDate() - 1)
        const lastWedKey = this.dateToKey(lastWed)
        for (let d = 1; d <= daysInMonth; d++) {
          const date = new Date(year, monthIdx, d)
          const dow = date.getDay()
          if (dow === 0) add(date, 'Pastor Paul', '', '#60a5fa')
          if (dow === 3) {
            if (this.dateToKey(date) === lastWedKey) {
              add(date, 'Small Groups', '10:00', '#10b981')
              add(date, 'Small Groups', '19:45', '#10b981')
            } else {
              add(date, 'Bible Study & PM', '19:30', '#6366f1')
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

/* Font family applied via Tailwind but we need to register them */
.font-sans { font-family: 'DM Sans', ui-sans-serif, system-ui, sans-serif; }
.font-mono { font-family: 'DM Mono', ui-monospace, monospace; }

/* ── Shared field primitives (used in multiple tabs) ── */
.field-input {
  background: #fff;
  border: 1px solid #e2e8f0;
  border-radius: 7px;
  padding: 6px 10px;
  color: #1e293b;
  font-family: 'DM Sans', sans-serif;
  font-size: 0.75rem;
  outline: none;
  transition: border-color 0.13s;
}
.field-input:focus { border-color: #1e293b; }
.field-input::placeholder { color: #94a3b8; }

.field-btn {
  background: #fff;
  border: 1px solid #e2e8f0;
  border-radius: 7px;
  padding: 6px 10px;
  color: #1e293b;
  font-family: 'DM Sans', sans-serif;
  font-size: 0.72rem;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.1s, border-color 0.1s;
}
.field-btn:hover { background: #f1f5f9; border-color: #1e293b; }
.field-btn:disabled { opacity: 0.3; cursor: not-allowed; }

/* ── Header action buttons ── */
.action-btn {
  background: #fff;
  border: 1px solid #e2e8f0;
  color: #334155;
  width: 30px;
  height: 30px;
  border-radius: 8px;
  cursor: pointer;
  font-size: 0.9rem;
  transition: background 0.1s, border-color 0.1s, color 0.1s;
}
.action-btn:hover { background: #f1f5f9; border-color: #1e293b; color: #1e293b; }

.hdr-btn {
  display: flex;
  align-items: center;
  gap: 5px;
  background: #fff;
  border: 1px solid #e2e8f0;
  color: #334155;
  padding: 0 13px;
  height: 30px;
  border-radius: 8px;
  cursor: pointer;
  font-family: 'DM Sans', sans-serif;
  font-size: 0.75rem;
  font-weight: 500;
  transition: background 0.1s, border-color 0.1s, color 0.1s;
  white-space: nowrap;
}
.hdr-btn:hover { background: #f1f5f9; border-color: #1e293b; color: #1e293b; }

/* ── Event chip: color-mix tint can't be expressed in Tailwind v3 ── */
.event-chip {
  background: color-mix(in srgb, var(--ec) 18%, #fff);
  border-left: 2px solid var(--ec);
}

/* ── Palette chip tints ── */
.palette-chip {
  background: color-mix(in srgb, var(--c) 10%, #fff);
  border: 1px solid color-mix(in srgb, var(--c) 30%, #e2e8f0);
}
.palette-chip:hover {
  background: color-mix(in srgb, var(--c) 18%, #fff);
  border-color: color-mix(in srgb, var(--c) 60%, #e2e8f0);
}
.p-label { color: color-mix(in srgb, var(--c) 70%, #1e293b); }

/* ── GCal template buttons ── */
.gcal-tmpl-btn {
  background: color-mix(in srgb, var(--c) 10%, #fff);
  border: 1px solid color-mix(in srgb, var(--c) 25%, #e2e8f0);
  color: color-mix(in srgb, var(--c) 70%, #1e293b);
}
.gcal-tmpl-btn:hover,
.gcal-tmpl-selected {
  background: color-mix(in srgb, var(--c) 20%, #fff);
  border-color: color-mix(in srgb, var(--c) 55%, #e2e8f0);
}

/* ── Today cell ring ── */
.today-cell::after {
  content: '';
  position: absolute;
  inset: 0;
  border: 1px solid #bfdbfe;
  pointer-events: none;
  border-radius: 1px;
}

/* ── Drag-over cell ── */
.drag-over-cell {
  background: #dbeafe !important;
  box-shadow: inset 0 0 0 2px #3b82f6;
}

/* ── School term bands ── */
.term-school { background: rgba(16, 185, 129, 0.25); }
.term-break  { background: rgba(245, 158, 11, 0.3); }

/* ── Modal transition ── */
.modal-enter-active, .modal-leave-active { transition: opacity 0.18s; }
.modal-enter-from, .modal-leave-to { opacity: 0; }
</style>