<template>
  <div class="view-analysis">
    <!-- Back Button -->
    <!-- <div class="head">
      <v-btn @click="()=>{
        let level = $route.params.level;
        // $router.replace(`${level}/our-services`)
      }">
        <span
          class="mdi mdi-chevron-left border-2 border-gray-900 dark:border-gray-300 rounded-xl py-2 px-4 text-[20px] cursor-pointer ms-3 hover:bg-gray-300 hover:text-gray-900 dark:hover:bg-gray-300 dark:hover:text-gray-900 duration-300"
        ></span>
      </v-btn>
    </div> -->

    <!-- File Upload Section -->
    <div class="import-file">
      <br />
      <div class="input">
        <div class="flex items-center justify-center w-full">
          <label
            for="dropzone-file"
            class="flex flex-col items-center justify-center duration-300 w-full h-64 border-2 border-gray-300 border-dashed rounded-lg cursor-pointer bg-gray-50 dark:hover:bg-gray-950 dark:bg-gray-900 hover:bg-gray-100 dark:border-gray-600 dark:hover:border-gray-500"
          >
            <div class="flex flex-col items-center justify-center pt-5 pb-6">
              <svg
                class="w-8 h-8 mb-4 text-gray-500 dark:text-gray-400"
                aria-hidden="true"
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 20 16"
              >
                <path
                  stroke="currentColor"
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  stroke-width="2"
                  d="M13 13h3a3 3 0 0 0 0-6h-.025A5.56 5.56 0 0 0 16 6.5 5.5 5.5 0 0 0 5.207 5.021C5.137 5.017 5.071 5 5 5a4 4 0 0 0 0 8h2.167M10 15V6m0 0L8 8m2-2 2 2"
                />
              </svg>
              <p class="mb-2 text-sm text-gray-700 dark:text-gray-400">
                <span class="font-semibold">Upload File</span>
              </p>
              <p class="text-xs text-gray-500 dark:text-gray-400">xlsx, xls</p>
            </div>
            <input
              @change="handleFileUpload"
              id="dropzone-file"
              type="file"
              title="Upload Service Table"
              accept=".xlsx, .xls"
              class="hidden"
            />
          </label>
        </div>
      </div>
    </div>

    <!-- Data Table Section -->
    <div class="tables">
      <div class="relative mx-4 overflow-x-auto shadow-md sm:rounded-lg">
        <table
          class="w-full text-sm text-left rtl:text-right text-gray-500 dark:text-gray-400"
        >
          <thead
            class="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400"
          >
            <tr>
              <th
                v-for="(header, index) in rowsData[0]"
                :key="index"
                scope="col"
                class="px-6 py-3"
              >
                {{ header }}
              </th>
            </tr>
          </thead>
          <tbody>
            <tr
              class="group odd:bg-white odd:dark:bg-gray-900 even:bg-gray-50 even:dark:bg-gray-800 border-b dark:border-gray-700"
              v-for="(row, rowIndex) in dataTable"
              :key="rowIndex"
            >
              <td
                v-for="(cell, cellIndex) in row"
                :key="cellIndex"
                class="px-6 py-4 relative"
                @mouseover="hoveredCell = { rowIndex, cellIndex }"
                @mouseleave="hoveredCell = null"
                :class="{
                  'bg-blue-100 dark:bg-blue-900 text-blue-700 dark:text-blue-300':
                    hoveredCell?.rowIndex === rowIndex &&
                    hoveredCell?.cellIndex === cellIndex,
                  'transition duration-300 ease-in-out': true,
                }"
              >
                <span>{{ handelCell(cell) }}</span>
                <div
                  v-if="
                    hoveredCell?.rowIndex === rowIndex &&
                    hoveredCell?.cellIndex === cellIndex &&
                    cell?.formula
                  "
                  class="absolute z-10 bg-gray-100 dark:bg-gray-800 text-gray-700 dark:text-gray-200 rounded-md p-2 text-xs shadow-lg transition-opacity duration-300 ease-in-out"
                  style="top: -40px; left: 50%; transform: translateX(-50%)"
                >
                  Formula: {{ cell.formula }}
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
    <span
      class="mdi mdi-chart-bell-curve bg-black dark:bg-white dark:text-black text-white rounded-md p-2 fixed -bottom-3 cursor-pointer left-12 text-[35px]"
      @click="showGraphic = true"
    ></span>
    <div
      class="graphs fixed duration-[1.5s] w-full min-h-screen"
      :class="showGraphic ? 'top-0' : 'top-full'"
    >
      <span
        class="mdi mdi-close-circle-outline absolute top-5 cursor-pointer right-8 text-[35px] text-black z-20"
        @click="showGraphic = false"
      ></span>
      <!-- Background Gradient -->
      <div class="background-gradient"></div>

      <!-- Sidebar Buttons -->
      <div class="sidebar">
        <button @click="showChart('interarrival')">
          Interarrival Histogram
        </button>
        <button @click="showChart('serviceBar')">Service - Bar</button>
        <button @click="showChart('servicePie')">Service - Pie</button>
        <button @click="showChart('systemBar')">System State - Bar</button>
        <button @click="showChart('systemPie')">System State - Pie</button>
        <button @click="showChart('customerStateBar')">
          Customer State - Bar
        </button>
        <button @click="showChart('customerStatePie')">
          Customer State - Pie
        </button>
        <button @click="showChart('customerTimeline')">
          Customer Timeline
        </button>
      </div>

      <!-- Chart Container -->
      <div class="chart-container">
        <div class="animated-arrows">&#8594; &#8594;</div>
        <canvas ref="chartCanvas"></canvas>
      </div>
    </div>
  </div>
</template>
<script setup>
import { ref, watchEffect } from 'vue'
import Chart from 'chart.js/auto'
import ExcelJS from 'exceljs'
import { useRoute } from 'vue-router'
let showGraphic = ref(false)
let level = useRoute().params.level
const rowsData = ref([])
const dataTable = ref([])
const hoveredCell = ref(null) // Track hovered cell

const handelCell = val => {
  return (val || val == 0) && typeof val !== 'object'
    ? val
    : val?.result != null
      ? val?.result
      : null
}

const handleFileUpload = async event => {
  const file = event.target.files[0]
  if (file) {
    const workbook = new ExcelJS.Workbook()
    const arrayBuffer = await file.arrayBuffer()
    await workbook.xlsx.load(arrayBuffer)

    const worksheet = workbook.worksheets[0] // Access the first sheet
    const startRow = level === 'Intermediate' ? 3 : 5 // Adjust starting row based on level
    const parsedRows = []

    worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
      const rowData = row.values.slice(1) // Exclude the first element (empty in ExcelJS)
      parsedRows.push(rowData)
      // if (rowIndex >= startRow) {
      // }
    })

    rowsData.value = parsedRows // Update the reactive state
    dataTable.value = parsedRows.slice(1) // Update the reactive state
    console.log(dataTable.value)
    getChart(dataTable.value)
  }
}
const getChart = data => {
  // get interarrival
  for (let i = 4; i < data.length; i++) {
    if (level == 'Beginner') {
      if (data[i][1]?.result != undefined) interarrival.push(data[i][1]?.result)
      if (data[i][2]?.result != undefined) arrivalTimes.push(data[i][2]?.result)
      if (data[i][4]?.result != undefined) services.push(data[i][4]?.result)
      if (data[i][7]?.result != undefined)
        endServiceTimes.push(data[i][7]?.result)
      if (data[i][8]?.result != undefined) systemState.push(data[i][8]?.result)
      if (data[i][9]?.result != undefined)
        customerState.push(data[i][9]?.result)
    } else {
      if (data[i][1]?.result != undefined) interarrival.push(data[i][2]?.result)
      if (data[i][2]?.result != undefined) arrivalTimes.push(data[i][3]?.result)
      if (data[i][4]?.result != undefined) services.push(data[i][6]?.result)
      if (data[i][7]?.result != undefined)
        endServiceTimes.push(data[i][7]?.result)
      if (data[i][8]?.result != undefined)
        systemState.push(data[i][9]?.result ? 'Idel' : 'Busy')
      if (data[i][9]?.result != undefined)
        customerState.push(data[i][8]?.result)
    }
  }
}
const chartCanvas = ref(null)
let chart = null

const interarrival = []
// Data for charts
const services = []
const systemState = []
const customerState = []
const arrivalTimes = [] //3
const endServiceTimes = [] //8

const showChart = type => {
  if (chart) chart.destroy() // Clear existing chart
  const ctx = chartCanvas.value.getContext('2d')
  if (!ctx) return
  let config = {}
  const container = document.querySelector('.chart-container')
  container.style.animation = 'none'
  setTimeout(() => (container.style.animation = ''), 0)
  if (type === 'interarrival') {
    const bins = Array(11).fill(0)
    interarrival.forEach(v => bins[v]++)
    const labels = Array.from({ length: 11 }, (_, i) => i)

    config = {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            label: 'Frequency of Interarrival Times',
            data: bins,
            backgroundColor: '#3498db',
          },
        ],
      },
    }
  } else if (type.includes('service')) {
    const counts = services.reduce(
      (acc, curr) => ((acc[curr] = (acc[curr] || 0) + 1), acc),
      {},
    )
    const labels = Object.keys(counts)
    const data = Object.values(counts)

    config = {
      type: type === 'serviceBar' ? 'bar' : 'pie',
      data: {
        labels,
        datasets: [
          {
            label: 'Service Count',
            data,
            backgroundColor: [
              '#1abc9c',
              '#e74c3c',
              '#9b59b6',
              '#f1c40f',
              '#2ecc71',
            ],
          },
        ],
      },
    }
  } else if (type.includes('system')) {
    const counts = systemState.reduce(
      (acc, curr) => ((acc[curr] = (acc[curr] || 0) + 1), acc),
      {},
    )
    const labels = Object.keys(counts)
    const data = Object.values(counts)

    config = {
      type: type === 'systemBar' ? 'bar' : 'pie',
      data: {
        labels,
        datasets: [
          {
            label: 'System State Count',
            data,
            backgroundColor: ['#3498db', '#e67e22'],
          },
        ],
      },
    }
  } else if (type.includes('customerState')) {
    const counts = customerState.reduce(
      (acc, curr) => ((acc[curr] = (acc[curr] || 0) + 1), acc),
      {},
    )
    const labels = Object.keys(counts)
    const data = Object.values(counts)

    config = {
      type: type === 'customerStateBar' ? 'bar' : 'pie',
      data: {
        labels,
        datasets: [
          {
            label: 'Customer State Count',
            data,
            backgroundColor: ['#3498db', '#e74c3c'],
          },
        ],
      },
    }
  } else if (type === 'customerTimeline') {
    let customers = 0
    const events = [
      ...arrivalTimes.map(t => ({ time: t, type: 'arrival' })),
      ...endServiceTimes.map(t => ({ time: t, type: 'departure' })),
    ].sort((a, b) => a.time - b.time)

    const times = []
    const counts = []
    events.forEach(event => {
      times.push(event.time)
      counts.push(customers)
      customers += event.type === 'arrival' ? 1 : -1
      times.push(event.time)
      counts.push(customers)
    })

    config = {
      type: 'line',
      data: {
        labels: times,
        datasets: [
          {
            label: 'Number of Customers Over Time',
            data: counts,
            borderColor: '#8e44ad',
            backgroundColor: 'rgba(142, 68, 173, 0.2)',
            fill: true,
            stepped: true,
          },
        ],
      },
      options: {
        responsive: true,
        plugins: { legend: { display: true } },
      },
    }
  }

  chart = new Chart(ctx, config)
}

watchEffect(() => {
  // Optional setup logic if you need reactivity
})
</script>

<style lang="scss">
body {
  overflow: auto !important;
}
/* Add hover animations */
.table-row:hover {
  transition:
    background-color 0.3s ease,
    color 0.3s ease;
}

td {
  position: relative;
  cursor: pointer;
}

td:hover {
  background-color: #f3f4f6; /* Light gray for light mode */
  color: #1e40af; /* Dark blue for light mode */
  transition: all 0.3s ease-in-out;
}

td:hover .tooltip {
  opacity: 1;
  transform: translateY(0);
}

/* Tooltip styling */
.tooltip {
  opacity: 0;
  position: absolute;
  z-index: 10;
  background-color: #374151; /* Dark gray for dark mode */
  color: white;
  padding: 5px 10px;
  border-radius: 5px;
  font-size: 0.75rem;
  white-space: nowrap;
  transition:
    opacity 0.2s ease,
    transform 0.2s ease;
  transform: translateY(10px);
}
.graphs {
  /* Sidebar Buttons */
  .sidebar {
    width: 230px;
    background-color: #2c3e50;
    padding: 20px;
    box-shadow: 2px 0 5px rgba(0, 0, 0, 0.1);
    display: flex;
    flex-direction: column;
    gap: 15px;
    height: 100vh;
  }
  button {
    background-color: #3498db;
    color: #ffffff;
    padding: 12px;
    border: none;
    cursor: pointer;
    border-radius: 5px;
    text-align: center;
    transition: 0.3s;
    font-weight: bold;
  }
  button:hover {
    background-color: #2980b9;
    transform: scale(1.05);
  }

  /* Chart Container */
  .chart-container {
    width: calc(100% - 230px);
    height: 100%;
    position: relative;
    background-color: white;
    box-shadow: -2px 0 5px rgba(0, 0, 0, 0.1);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-direction: column;
    opacity: 0;
    transform: translateX(100%);
    animation: slideIn 1.5s forwards ease-out;
    margin-inline-start: 230px;
    right: 0;
    top: -748px;
    height: 100vh;
  }
  canvas {
    width: 90% !important;
    max-height: 90%;
    margin: 10px auto;
  }

  /* Animations */
  @keyframes slideIn {
    from {
      opacity: 0;
      transform: translateX(100%);
    }
    to {
      opacity: 1;
      transform: translateX(0);
    }
  }

  /* Animated Arrows */
  .animated-arrows {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    font-size: 30px;
    animation: moveArrows 1.5s forwards ease-out;
  }
  @keyframes moveArrows {
    0% {
      left: 50%;
      opacity: 1;
    }
    100% {
      left: 60%;
      opacity: 0;
    }
  }

  /* Animated Background Radial Gradient */
  .background-gradient {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: radial-gradient(circle, #3498db, #1abc9c, #9b59b6);
    animation: moveBackground 5s linear infinite;
    z-index: -1;
  }

  @keyframes moveBackground {
    0% {
      background-position: center;
    }
    50% {
      background-position: 60% 60%;
    }
    100% {
      background-position: center;
    }
  }
}
</style>
