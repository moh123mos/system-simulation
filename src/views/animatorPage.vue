<template>
  <div id="app-animation">
    <div class="w-full content">
      <h1
        class="text-3xl font-bold mb-6 text-gray-800 dark:text-gray-300 text-center"
      >
        Queueing System Animator
      </h1>

      <input
        @change="handleFileUpload"
        id="dropzone-file"
        type="file"
        title="Upload Service Table"
        accept=".xlsx, .xls"
        class=""
      />

      <!-- Info Section -->
      <div class="flex justify-between mb-6 text-lg gap-24">
        <div>
          Number of customers in queue:
          <span class="font-bold text-blue-500">{{ queue.length -1}}</span>
        </div>
        <div>
          System state:
          <span class="font-bold text-green-500">{{
            systemBusy ? 'Busy' : 'Idle'
          }}</span>
        </div>
        <div>
          Time: <span class="font-bold text-red-500">{{ time }}</span
          >s
        </div>
      </div>

      <!-- Simulation Area -->
      <div
        class="relative w-[90%] bg-white dark:bg-black border rounded-lg h-56 flex items-center overflow-hidden shadow-lg"
      >
        <div class="queue ml-4">
          <div
            v-for="(_, index) in queue"
            :key="index"
            class="person"
            :class="{ first: index == queue.length - 1 }"
          ></div>
        </div>
        <div class="system">
          <div
            v-if="processingCustomer"
            class="person"
            style="animation: fadeOut 0.5s forwards"
          ></div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import ExcelJS from 'exceljs'
const customers = ref([
  { arrivalTime: 1, serviceDuration: 5 },
  { arrivalTime: 3, serviceDuration: 3 },
  { arrivalTime: 4, serviceDuration: 4 },
  { arrivalTime: 6, serviceDuration: 2 },
  { arrivalTime: 7, serviceDuration: 6 },
  { arrivalTime: 1, serviceDuration: 5 },
  { arrivalTime: 3, serviceDuration: 3 },
  { arrivalTime: 4, serviceDuration: 4 },
  { arrivalTime: 6, serviceDuration: 2 },
  { arrivalTime: 7, serviceDuration: 6 },
  { arrivalTime: 1, serviceDuration: 5 },
  { arrivalTime: 3, serviceDuration: 3 },
  { arrivalTime: 4, serviceDuration: 4 },
  { arrivalTime: 6, serviceDuration: 2 },
  { arrivalTime: 7, serviceDuration: 6 },
  { arrivalTime: 1, serviceDuration: 5 },
  { arrivalTime: 3, serviceDuration: 3 },
  { arrivalTime: 4, serviceDuration: 4 },
  { arrivalTime: 6, serviceDuration: 2 },
  { arrivalTime: 7, serviceDuration: 6 },
  { arrivalTime: 1, serviceDuration: 5 },

])
const time = ref(0)
const queue = ref([])
const systemBusy = ref(false)
const processingCustomer = ref(null)
let totalTime = ref(0)
const processCustomer = customer => {
  systemBusy.value = true
  processingCustomer.value = customer
  setTimeout(() => {
    systemBusy.value = false
    processingCustomer.value = null
    if (queue.value.length > 0) {
      // document.querySelector('.person.first').style = "animation: moveToSystem 0.5s forwards"
      const nextCustomer = queue.value.shift()
      processCustomer(nextCustomer)
    }
  }, customer.serviceDuration * 1000)
}

const startAnime = () => {
  const simulation = setInterval(() => {
    time.value++
    customers.value.forEach(customer => {
      if (customer.arrivalTime === time.value) {
        if (!systemBusy.value) {
          processCustomer(customer)
        } else {
          queue.value.push(customer)
          // addToQueue()
        }
      }
    })

    if (time.value > totalTime.value)
      clearInterval(simulation), location.reload()
  }, 1000)
}

const handleFileUpload = async event => {
  const file = event.target.files[0]
  if (file) {
    const workbook = new ExcelJS.Workbook()
    const arrayBuffer = await file.arrayBuffer()
    await workbook.xlsx.load(arrayBuffer)

    const worksheet = workbook.worksheets[0] // Access the first sheet
    // const startRow = level === 'Intermediate' ? 3 : 5 // Adjust starting row based on level
    const parsedRows = []

    worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
      const rowData = row.values.slice(1) // Exclude the first element (empty in ExcelJS)
      parsedRows.push(rowData)
      // if (rowIndex >= startRow) {
      // }
    })

    const rowsData = parsedRows // Update the reactive state
    let dataTable = parsedRows.slice(1) // Update the reactive state
    console.log(parsedRows)
    getCustomersData(parsedRows.slice(1))
    startAnime()
  }
}
const getCustomersData = data => {
  // get interarrival
  let time = {};
  // customers.value = [];
  console.log(data);
  for (let i = 4; i < data?.length; i++) {
    if (data[i][2]?.result != undefined)
      time['arrivalTime'] = data[i][2]?.result
    if (data[i][6]?.result != undefined)
      time['serviceDuration'] = data[i][6]?.result
    // customers.value.push(time);
    console.log(time);
  }
  totalTime.value =
    +customers.value[customers.value.length - 1]?.arrivalTime +
    +customers.value[customers.value.length - 1]?.serviceDuration + 65
  console.log(totalTime.value)
}
</script>

<style lang="scss">
#app-animation {
  .content {
    display: flex;
    flex-direction: column;
    justify-content: space-evenly;
    align-items: center;
    width: 100%;
    height: calc(100vh - 80px);
  }
  .person {
    width: 60px;
    height: 60px;
    background-image: url('https://i.pravatar.cc/60'); /* Placeholder image */
    background-size: cover;
    background-position: center;
    border: 3px solid #4a5568;
    border-radius: 50%;
    margin-right: 8px;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2);
    animation: fadeIn 0.5s ease-in;
    &.first {
      animation: moveToSystem 3s ease;
    }
  }

  /* Queue Area */
  .queue {
    display: flex;
    align-items: center;
    justify-content: start;
    overflow: hidden;
    padding: 12px;
    background: #edf2f7;
    border: 2px solid #cbd5e0;
    border-radius: 8px;
    height: 72px;
    flex-grow: 1;
    box-shadow: inset 0 2px 6px rgba(0, 0, 0, 0.1);
  }

  /* System Styling */
  .system {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 120px;
    height: 100%;
    background-color: #c3dafe;
    border: 2px solid #90cdf4;
    border-radius: 12px;
    margin-left: 16px;
    flex-shrink: 0;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    background-image: url('https://cdn-icons-png.flaticon.com/512/3135/3135768.png'); /* System Image */
    background-size: 50%;
    background-repeat: no-repeat;
    background-position: center;
  }

  /* Animations */
  @keyframes moveToSystem {
    0% {
      transform: translateX(0);
      opacity: 1;
    }
    100% {
      transform: translateX(100vw);
      opacity: 1;
    }
  }

  @keyframes fadeOut {
    0% {
      opacity: 1;
      transform: scale(1);
    }
    100% {
      opacity: 0;
      transform: scale(0.5);
    }
  }

  @keyframes fadeIn {
    from {
      opacity: 0;
      transform: scale(0.8);
    }
    to {
      opacity: 1;
      transform: scale(1);
    }
  }
}
</style>
