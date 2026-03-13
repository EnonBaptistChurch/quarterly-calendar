<template>
  <div class="relative inline-block w-full mb-10" @keydown.prevent.stop="handleKeydown" tabindex="0" ref="wrapper">
    <!-- Input field -->
    <div @click="toggleDropdown"
         class="cursor-pointer border border-gray-300 rounded-md bg-white px-3 py-2  flex justify-between items-center shadow-sm hover:shadow-md transition w-full">
      <span>{{ modelValue || 'HH:mm' }}</span>
      <svg class="w-4 h-4 text-gray-500 transition-transform duration-200"
           :class="{'rotate-180': open}"
           fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"/>
      </svg>
    </div>

    <!-- Dropdown -->
    <div v-if="open"
         class="absolute z-10 mt-1 bg-white border border-gray-300 rounded-md max-h-[90px] overflow-auto shadow-lg"
         :style="{ width: dropdownWidth + 'px' }"
         ref="dropdown">
      <div v-for="(time, index) in times" :key="time"
           :ref="el => timeRefs[index] = el"
           @click="selectTime(time)"
           :class="[
             'px-3 py-1 cursor-pointer',
             highlightedIndex === index ? 'bg-blue-100' : '',
           ]">
        {{ time }}
      </div>
    </div>
  </div>
</template>

<script>
export default {
  name: 'TimePicker',
  props: { modelValue: String },
  data() {
    return {
      open: false,
      highlightedIndex: -1,
      times: [],
      timeRefs: [],
      dropdownWidth: 0,
    }
  },
  mounted() {
    // Generate all 15-min interval times
    const list = [];
    for (let h = 7; h < 24; h++) {
      for (let m of [0,15,30,45]) {
        list.push(`${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}`);
      }
    }
    this.times = list;

    document.addEventListener('click', this.handleClickOutside);

    // Set initial dropdown width
    this.$nextTick(() => {
      if (this.$refs.wrapper) {
        this.dropdownWidth = this.$refs.wrapper.offsetWidth;
      }
    });

    window.addEventListener('resize', this.updateDropdownWidth);
  },
  beforeUnmount() {
    document.removeEventListener('click', this.handleClickOutside);
    window.removeEventListener('resize', this.updateDropdownWidth);
  },
  methods: {
    toggleDropdown() {
      this.open = !this.open;
      if (this.open) this.scrollToSelected();
    },
    selectTime(time) {
      this.$emit('update:modelValue', time);
      this.open = false;
    },
    handleClickOutside(e) {
      if (!this.$el.contains(e.target)) this.open = false;
    },
    handleKeydown(e) {
      if (!this.open) {
        if (e.key === 'ArrowDown' || e.key === 'Enter') {
          this.open = true;
          this.highlightedIndex = this.times.indexOf(this.modelValue) || 0;
          this.scrollToSelected();
        }
        return;
      }

      if (e.key === 'ArrowDown') {
        this.highlightedIndex = (this.highlightedIndex + 1) % this.times.length;
        this.scrollToHighlighted();
      } else if (e.key === 'ArrowUp') {
        this.highlightedIndex = (this.highlightedIndex - 1 + this.times.length) % this.times.length;
        this.scrollToHighlighted();
      } else if (e.key === 'Enter') {
        if (this.highlightedIndex >= 0) this.selectTime(this.times[this.highlightedIndex]);
      } else if (e.key === 'Escape') {
        this.open = false;
      }
    },
    scrollToSelected() {
      this.highlightedIndex = this.times.indexOf(this.modelValue) || 0;
      this.$nextTick(() => this.scrollToHighlighted());
    },
    scrollToHighlighted() {
      const el = this.timeRefs[this.highlightedIndex];
      if (el) el.scrollIntoView({ block: 'nearest' });
    },
    updateDropdownWidth() {
      if (this.$refs.wrapper) {
        this.dropdownWidth = this.$refs.wrapper.offsetWidth;
      }
    }
  }
}
</script>