<template>
  <div class="hero">
    <v-container>
      <v-btn class="go-back"
        ><router-link to="/"
          ><v-icon>mdi-chevron-left</v-icon></router-link
        ></v-btn
      >
      <div class="head">
        <h1 class="text-center">Simulation Setup</h1>
        <v-btn class="dark-light-mode-btn" @click="toggleTheme"
          ><v-icon>{{
            isDarkMode ? 'mdi-white-balance-sunny' : 'mdi-weather-night'
          }}</v-icon></v-btn
        >
      </div>
      <v-form>
        <div class="inputs">
          <div class="no-customers">
            <v-text-field
              label="Number Of Customers"
              type="number"
              v-model="numberValue"
              :min="1"
            ></v-text-field>
          </div>
          <div v-if="(level==='Beginner')" class="min-inter-arrival">
            <v-text-field
              label="Min Interarrival "
              type="number"
              v-model="numberValue"
              :min="1"
            ></v-text-field>
          </div>
          <div v-if="(level==='Beginner')" class="max-inter-arrival">
            <v-text-field
              label="Max Interarrival"
              type="number"
              v-model="numberValue"
              :min="1"
            ></v-text-field>
          </div>
          <div class="upload-service">
            <label for="upload-file">{{ level==='Beginner'?'Upload Service Table': 'Upload Data Table' }}</label>
            <input id="upload-file" type="file" title="Upload Service Table">
          </div>
        </div>
        <div class="text-center">
          <v-btn  type="submit" style="text-transform: capitalize"> Simulate and Save </v-btn>
        </div>
      </v-form>
    </v-container>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import { useRoute } from 'vue-router';
import { useTheme } from 'vuetify'
const theme = useTheme()
// Define a computed property for the current theme name
const isDarkMode = computed(() => theme.global.name.value === 'dark')
// Toggle theme between light and dark
const toggleTheme = () => {
  theme.global.name.value = isDarkMode.value ? 'light' : 'dark'
}
const route = useRoute()
let level = route.params.level;
</script>

<style lang="scss" scoped>
a {
  text-decoration: none;
  color: inherit;
}
.hero {
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100vh;
  .link {
  }
  .go-back {
    position: fixed;
    left: 20px;
    top: 20px;
    font-size: 25px;
  }
  .head {
    position: relative;
    font-size: 30px;
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
  .upload-service {
    display: flex;
    flex-direction: column;
    margin-bottom: 20px;
  }
}
</style>
