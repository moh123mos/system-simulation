<template>
  <div
    class="container m-auto py-5 px-4 flex justify-between items-center nav-bar"
  >
    <div
      class="logo text-[24px] cursor-pointer"
      @click="
        () => {
          $router.replace('/')
        }
      "
    >
      zootopia
    </div>
    <ul class="flex justify-between items-center gap-9">
      <!-- <li v-for="link in links" :key="link">
        <router-link :to="link.url" class="text-[18px] duration-300 hover:text-[#00F0FF]">{{
          link.title
        }}</router-link>
      </li> -->
    </ul>
    <ul class="flex justify-between items-center gap-2">
      <li class="cursor-pointer">
        <span
          class="text-[20px] mdi"
          @click="handleToggleDarkMode"
          :class="
            isDarkMode.status ? 'mdi-white-balance-sunny' : 'mdi-weather-night'
          "
        ></span>
      </li>
      <li v-for="icon in icons" :key="icon">
        <router-link
          :to="icon.url"
          class="duration-300 hover:text-[#00F0FF] text-[20px]"
        >
          <span :class="icon.mdi"></span>
        </router-link>
      </li>
    </ul>
  </div>
</template>
<script setup>
import { ref } from 'vue'

const links = ref([
  {
    title: 'Smart Contracts',
    url: '',
  },
  {
    title: 'Services',
    url: '',
  },
  {
    title: 'Solutions',
    url: '',
  },
  {
    title: 'Roadmap',
    url: '',
  },
  {
    title: 'Whitepaper',
    url: '',
  },
])
const icons = ref([
  {
    title: 'github',
    mdi: 'mdi mdi-github',
    url: '',
  },
  {
    title: 'reddit',
    mdi: 'mdi mdi-reddit',
    url: '',
  },
  {
    title: 'linkedin',
    mdi: 'mdi mdi-linkedin',
    url: '',
  },
  {
    title: 'twitter',
    mdi: 'mdi mdi-twitter',
    url: '',
  },
])

// toggle dark mode
import { onMounted } from 'vue'
import toggleDarkMode from '@/composables/useToggleMode'
const isDarkMode = ref({
  status: true,
})
onMounted(() => {
  const savedTheme = localStorage.getItem('theme')
  if (savedTheme === 'dark') {
    document.documentElement.classList.add('dark')
    isDarkMode.value.status = true
  } else {
    document.documentElement.classList.remove('dark')
    isDarkMode.value.status = false
  }
})
const handleToggleDarkMode = () => {
  toggleDarkMode(isDarkMode.value)
}
</script>
<style lang="scss">
.nav-bar {
  .logo {
    font-family: 'Chillax', sans-serif;
  }
}
</style>
