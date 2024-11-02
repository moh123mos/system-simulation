import { createApp } from 'vue'
import { createPinia } from 'pinia'

import App from './App.vue'
import router from './router'
// Vuetify
import 'vuetify/styles'
import { createVuetify } from 'vuetify'
import * as components from 'vuetify/components'
import * as directives from 'vuetify/directives'
import { aliases, mdi } from 'vuetify/iconsets/mdi';
import '@mdi/font/css/materialdesignicons.css';
const vuetify = createVuetify({
  components,
  directives,
  theme: {
    defaultTheme: 'dark',
    themes: {
      light: {
        dark: false,
        colors: {
          background: '#FFFFFF',
          surface: '#FFFFFF',
          primary: '#6200EE',
          secondary: '#03DAC6',
          accent: '#03DAC6',
        },
      },
      dark: {
        dark: true,
        colors: {
          background: '#1d2125',
          surface: '#1d2125',
          primary: '#BB86FC',
          secondary: '#03DAC6',
          accent: '#03DAC6',
        },
      },
    },
  },
  icons: {
    defaultSet: 'mdi',
    aliases,
    sets: { mdi },
  },
})

const app = createApp(App)

app.use(createPinia())
app.use(router)
app.use(vuetify)

app.mount('#app')
