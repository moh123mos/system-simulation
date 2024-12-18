<template>
  <div class="our-services pt-14">
    <div class="cards">
      <div
        class="card duration-300 bg-white dark:bg-black text-black dark:text-white border-black dark:border-white pb-4"
        v-for="(card, i) in cards"
        :key="i"
      >
        <img :src="card.image" alt="image" class="h-[245px]" />
        <h3 class="card-title">{{ card.title }}</h3>
        <P class="card-description">
          {{ card.description }}
        </P>
        <button class="btn cursor-pointer">
          <a
            v-if="card.url === 'create-analysis' && level === 'Beginner'"
            href="http://localhost:5173/Beginner/create-analysis"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'create-analysis' && level === 'Intermediate'"
            href="http://localhost:5173/Intermediate/create-analysis"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'create-analysis' && level === 'Advanced'"
            href="http://localhost:5173/Advanced/create-analysis"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'view-analysis' && level === 'Beginner'"
            href="http://localhost:5173/Beginner/view-analysis"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'view-analysis' && level === 'Intermediate'"
            href="http://localhost:5173/Intermediate/view-analysis"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'view-analysis' && level === 'Advanced'"
            href="http://localhost:5173/Advanced/view-analysis"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'animator' && level === 'Beginner'"
            href="http://localhost:5173/Beginner/animator"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'animator' && level === 'Intermediate'"
            href="http://localhost:5173/Intermediate/animator"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'animator' && level === 'Advanced'"
            href="http://localhost:5173/Advanced/animator"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'ai-analyzer' && level === 'Beginner'"
            href="http://localhost:5173/Beginner/ai-analyzer"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'ai-analyzer' && level === 'Intermediate'"
            href="http://localhost:5173/Intermediate/ai-analyzer"
          >
            {{ card.btn }}
          </a>
          <a
            v-if="card.url === 'ai-analyzer' && level === 'Advanced'"
            href="http://localhost:5173/Advanced/ai-analyzer"
          >
            {{ card.btn }}
          </a>
        </button>
      </div>
    </div>
    <!-- <img src ="@/assets/images/" > -->
  </div>
</template>
<script setup>
import { ref } from 'vue'
import { useRoute, useRouter } from 'vue-router'
const route = useRoute()
const router = useRouter()
let level = route.params.level
const cards = ref([
  {
    image: '/simulator.jpg',
    title: 'Simulator Master',
    description: `Dive into immersive simulations, mastering challenges with precision.
Explore realistic scenarios to enhance your skills and decision-making.`,
    btn: 'Simulator',
    url: `create-analysis`,
  },
  {
    image: '/visualizer.jpeg',
    title: ' Dynamic Visualiser',
    description: `Transform complex data into clear, interactive visuals effortlessly.
Explore insights with powerful tools for enhanced understanding.`,
    btn: 'Visualiser',
    url: `view-analysis`,
  },
  {
    image: '/zootopia.jpg',
    title: 'Creative Animator',
    description: `Bring ideas to life with stunning animations and smooth transitions.
Unleash your creativity to craft engaging and dynamic visual stories.`,
    btn: 'Animator',
    url: `animator`,
  },
  {
    image: '/ai_analyzer.jpeg',
    title: ' AI Analyzer',
    description: `Harness the power of AI to process and interpret complex data.
Gain actionable insights with speed, accuracy, and efficiency.`,
    btn: 'Ai analyzer',
    url: `ai-analyzer`,
  },
])
function navigateTo(url) {
  router.replace(url) // Replace the current URL
}
</script>
<style lang="scss">
.cards {
  display: flex;
  justify-content: space-evenly;
  margin-top: 60px;
  gap: 30px;
  max-width: 1500px;
  margin: auto;
}

.card {
  backdrop-filter: blur(10px); /*blur the background*/
  border: 2.5px solid rgba(255, 252, 252, 0.125);
  display: flex;
  flex-direction: column;
  justify-content: space-evenly;
  word-wrap: break-word;
  transition:
    transform 0.3s,
    box-shadow 0.3s;
  width: 300px;
  padding: 10px;
  gap: 15px;
  animation: borderAnimation 4s linear infinite;
}

@keyframes borderAnimation {
  0% {
    border-color: #ff3bff;
  }
  25% {
    border-color: #ecbfbf;
  }
  50% {
    border-color: #5c24ff;
  }
  75% {
    border-color: #d94fd5;
  }
  100% {
    border-color: #ff3bff;
  }
}

.card:hover {
  transform: scale(1.05);
  box-shadow:
    0 10px 20px hsla(0, 0%, 90.6%, 0.836),
    0 4px 8px hsla(0, 0%, 76.9%, 0.808);
}

/*.card {
    border-radius: 4px;
    background: #fff;
*/
.cards .card .card-title {
  color: #fff;
}

.cards .card .card-description {
  color: #808080;
}

.btn {
  display: block;
  margin: auto;
  background-color: #d94fd5;
  margin-top: 10px;
  padding: 10px 40px;
  border-radius: 50px;
  font-size: 20px;
  letter-spacing: 2px;
  background-color: transparent;
  border: 2px solid #d94fd5;
  position: relative;
  // text-align: center;
  //margin-bottom: -40px;
  animation: borderAnimation 4s linear infinite;

  @keyframes borderAnimation {
    0% {
      border-color: #ff3bff;
    }
    25% {
      border-color: #ecbfbf;
    }
    50% {
      border-color: #5c24ff;
    }
    75% {
      border-color: #d94fd5;
    }
    100% {
      border-color: #ff3bff;
    }
  }
  &:hover {
  }
}
</style>
