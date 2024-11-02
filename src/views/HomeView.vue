<template>
  <div class="hero">
    <v-container>
      <div class="head">
        <h1 class="text-center">Queueing System Simulator</h1>
        <v-btn class="dark-light-mode-btn" @click="toggleTheme"><v-icon>{{ isDarkMode ? 'mdi-white-balance-sunny' : 'mdi-weather-night' }}</v-icon></v-btn>
      </div>
      <main>
        <v-select
          clearable
          label="Select"
          :items="['Beginner', 'Intermediate', 'Advanced']"
          variant="solo-filled"
          v-model="level"
          value=""
        ></v-select>
        <div class="operations">
          <v-btn
            ><router-link class="link" :to="level+'/create-analysis/'"
              >Create Analysis</router-link
            ></v-btn
          >
          <v-btn
            ><router-link class="link" :to="level+'/view-analysis'"
              >Veiw Analysis</router-link
            ></v-btn
          >
        </div>
      </main>
    </v-container>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'
import { useTheme } from 'vuetify'
let level = ref('Beginner')
const theme = useTheme()
// Define a computed property for the current theme name
const isDarkMode = computed(() => theme.global.name.value === 'dark')
// Toggle theme between light and dark
const toggleTheme = () => {
  theme.global.name.value = isDarkMode.value ? 'light' : 'dark'
}
</script>

<style lang="scss">
.hero {
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100vh;
  .link {
    text-decoration: none;
    color: inherit;
  }
  .head {
    position: relative;
    margin-bottom: 30px;
    .dark-light-mode-btn {
      position: fixed;
      right: 20px;
      top: 20px;
    }
  }
  .operations {
    display: flex;
    justify-content: space-around;
  }
}
</style>
