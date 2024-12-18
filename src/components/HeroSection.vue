<template>
  <div class="items flex-col justify-start min-h-[calc(100vh-80px)]">
    <div class="info">
      <h1 class="main-title">Ai powerd Simulator</h1>
      <h1>Smart analysis</h1>
      <p>
        The AI Simulator creates and displays detailed simulation tables based
        on user inputs. <br />
        It also integrates an AI model that answers questions about the data,
        <br />
        providing insights and analysis.
      </p>
      <select v-model="level">
        <option value="Beginner">Single Server QS Random</option>
        <option value="Intermediate">Single Server QS probability</option>
        <option value="Advanced">Double Sever</option>
      </select>
      <button>
        <router-link class="link" :to="level + '/our-services/'"
          >start</router-link
        >
      </button>
    </div>
    <!-- info -->
    <div class="ball ball1"></div>
    <div class="ball ball2"></div>
    <div class="ball ball3"></div>
  </div>
  <!-- items -->
</template>
<style lang="scss">
body,
html {
  width: 100%;
  height: 100%;
  overflow: hidden;
  margin: 0;
  padding: 0;
  position: relative;
  background-color: black;
}

* {
  padding: 0;
  margin: 0;
  box-sizing: border-box;
}

.info {
  margin-top: 25px;
}

.items {
  background-image: url(@/assets/images/robots.png);
  background-size: contain;
  background-position: center bottom;
  background-repeat: no-repeat;
  text-align: center;
  z-index: 1;
  animation: gradientBackground 10s linear infinite;
}

.items img {
  margin-bottom: -110px;
  width: 100%;
  height: 100%;
}

.items p {
  letter-spacing: 2px;
  margin-bottom: 10px;
}

.items select {
  font-size: 20px;
  letter-spacing: 2px;
  border: 2px solid #d94fd5;
  padding: 16px 150px;
  border-radius: 50px;
  background-color: transparent;
  position: relative;
  animation: borderAnimation 4s linear infinite;
  appearance: none;
}

.items option {
  background-color: rgb(0, 0, 0);
  color: rgb(255, 255, 255);
  text-align: center;
  padding: 16px 150px;
  border-radius: 50px;
}

.items h1 {
  font-size: 60px;
  font-weight: bold;
  letter-spacing: 2px;
}

.items .main-title {
  margin-bottom: -20px;
  background: linear-gradient(45deg, #ff3bff, #ecbfbf, #5c24ff, #d94fd5);
  background-size: 400% 400%;
  -webkit-background-clip: text;
  color: transparent;
  animation: gradientAnimation 5s ease infinite;
}

@keyframes gradientAnimation {
  0% {
    background-position: 0% 50%;
  }
  50% {
    background-position: 100% 50%;
  }
  100% {
    background-position: 0% 50%;
  }
}

.items button {
  display: block;
  margin: auto;
  background-color: #d94fd5;
  margin-top: 10px;
  padding: 13px 40px;
  border-radius: 50px;
  font-size: 20px;
  letter-spacing: 2px;
  background-color: transparent;
  border: 2px solid #d94fd5;
  position: relative;
  margin-bottom: -40px;
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
// =======================================
.ball {
  width: 0.4px;
  height: 0.4px;
  border-radius: 50%;
  backdrop-filter: blur(10px);
  box-shadow: 0px 0px 70px 40px rgba(217, 79, 213, 0.8);
  position: absolute;
  animation:
    moveBall 8s ease-in-out infinite,
    changeColor 5s infinite;
  z-index: 0;
}

.ball1 {
  background-color: #ff3bff;
  animation:
    moveBall1 8s ease-in-out infinite,
    changeColor 5s infinite;
}

.ball2 {
  background-color: #5c24ff;
  animation:
    moveBall2 10s ease-in-out infinite,
    changeColor 5s infinite;
}

.ball3 {
  background-color: #d94fd5;
  animation:
    moveBall3 12s ease-in-out infinite,
    changeColor 5s infinite;
}

@keyframes moveBall1 {
  0% {
    top: 0%;
    left: 0%;
  }
  25% {
    top: 10%;
    left: 80%;
  }
  50% {
    top: 80%;
    left: 50%;
  }
  75% {
    top: 20%;
    left: 20%;
  }
  100% {
    top: 0%;
    left: 0%;
  }
}

@keyframes moveBall2 {
  0% {
    top: 0%;
    left: 100%;
  }
  25% {
    top: 50%;
    left: 80%;
  }
  50% {
    top: 50%;
    left: 20%;
  }
  75% {
    top: 70%;
    left: 70%;
  }
  100% {
    top: 0%;
    left: 100%;
  }
}

@keyframes moveBall3 {
  0% {
    top: 100%;
    left: 0%;
  }
  25% {
    top: 30%;
    left: 30%;
  }
  50% {
    top: 70%;
    left: 80%;
  }
  75% {
    top: 10%;
    left: 60%;
  }
  100% {
    top: 100%;
    left: 0%;
  }
}

@keyframes changeColor {
  0% {
    background-color: #ff3bff;
  }
  25% {
    background-color: #5c24ff;
  }
  50% {
    background-color: #5c24ff;
  }
  75% {
    background-color: #ff3bff;
  }
  100% {
    background-color: #5c24ff;
  }
}
</style>

<script setup>
import { ref } from 'vue';

let level = ref('Beginner');
</script>
