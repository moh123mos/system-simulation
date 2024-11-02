<script setup>
import { computed } from 'vue'
import { useTheme } from 'vuetify'

const theme = useTheme()

// Define a computed property for the current theme name
const isDarkMode = computed(() => theme.global.name.value === 'dark')

// Toggle theme between light and dark
// eslint-disable-next-line no-unused-vars
const toggleTheme = () => {
  theme.global.name.value = isDarkMode.value ? 'light' : 'dark'
}
</script>

<template>
  <div class="hero">
    <v-container>
      <div class="head">
        <h1>Queueing System Simulator</h1>
      </div>
      <main>
        <v-select
          clearable
          label="Select"
          :items="['Beginner', 'Intermediate', 'Advanced']"
          variant="solo-filled"
        ></v-select>
        <div class="operations d-flex justify-content-between">
          <v-btn
            ><router-link class="link" to="/create-analysis"
              >Create Analysis</router-link
            ></v-btn
          >
          <v-btn
            ><router-link class="link" to="/create-analysis"
              >Veiw Analysis</router-link
            ></v-btn
          >
        </div>
      </main>
    </v-container>
  </div>
</template>
<style lang="scss">
.hero {
  .link {
    text-decoration: none;
    color: inherit;
  }
  .operations {
    display: flex;
    justify-content: space-between;
  }
}
</style>
