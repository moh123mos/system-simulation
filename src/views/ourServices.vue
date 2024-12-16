<template>
  <div class="our-services">
    <div class="cards">
      <div
        class="card duration-300 bg-white dark:bg-black text-black dark:text-white border-black dark:border-white"
      >
        <img src="@/assets/images/simulator.jpg" alt="" />
        <h3 class="card-title">Simulator Master</h3>
        <P class="card-description">
          Dive into immersive simulations, mastering challenges with precision.
          Explore realistic scenarios to enhance your skills and
          decision-making.
        </P>
        <button class="btn">Simulate</button>
      </div>
    </div>
    <!-- <img src ="@/assets/images/" > -->
  </div>
</template>

<style lang="scss">
.cards {
  display: flex;
  justify-content: space-evenly;
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
  cursor: pointer;
  width: calc(90% / 3);
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
  cursor: pointer;
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

/*button {

//     } */
// .cards .card .btn {
//   width: fit-content;
//   padding: 5px;
//   margin: 2px;
// }

.btn {
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
