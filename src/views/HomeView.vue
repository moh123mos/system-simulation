<template>
  <div class="">
    <HeroSection />
  </div>
</template>

<script setup>
import HeroSection from '@/components/HeroSection.vue';
// import { ref } from 'vue'
// let level = ref('Beginner'
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
