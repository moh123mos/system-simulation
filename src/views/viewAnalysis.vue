<template>
  <div class="view-analysis">
    <v-btn class="dark-light-mode-btn" @click="toggleTheme"
      ><v-icon>{{
        isDarkMode ? 'mdi-white-balance-sunny' : 'mdi-weather-night'
      }}</v-icon></v-btn
    >
    <v-container>
      <div class="head">
        <v-btn class="go-back" @click="$router.push('/')"
          ><v-icon>mdi-chevron-left</v-icon></v-btn
        >
      </div>
      <main>
        <div class="import-file">
          <v-label color="primary" for="upload-file">Upload Excel File</v-label>
          <br />
          <input
            @change="handleFileUpload"
            id="upload-file"
            type="file"
            title="Upload Service Table"
            accept=".xlsx, .xls"
          />
        </div>
        <div class="tables">
          <v-layout class="table">
            <v-navigation-drawer v-model="drawer" permanent>
              <v-list-item>
                <h3>Simulation Tables</h3>
                <!-- <h3>Simulation Dashboard</h3> -->
              </v-list-item>

              <v-divider></v-divider>

              <v-list v-if="level === 'Beginner'" density="compact" nav>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Simulation Table"
                  value="simulation"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Service Table"
                  value="service"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Analysis Table"
                  value="analysis"
                ></v-list-item>
              </v-list>
              <v-list
                v-else-if="level === 'Intermediate'"
                density="compact"
                nav
              >
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Data Table"
                  value="data-table"
                  @click="tableSelected = 'data-table'"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Arrival probability Table"
                  value="arrival-probability"
                  @click="tableSelected = 'arrival-probability'"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Service probability Table"
                  value="service-probability"
                  @click="tableSelected = 'service-probability'"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Simulation Table"
                  value="simulation"
                  @click="tableSelected = 'simulation'"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="System analysis Table"
                  value="system-analysis"
                  @click="tableSelected = 'system-analysis'"
                ></v-list-item>
              </v-list>
              <v-list v-else density="compact" nav>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Data Table"
                  value="data"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Arrival probability Table"
                  value="arrival-probability"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Server 1 Table"
                  value="service-probability"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Server 2 Table"
                  value="service-probability"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="Simulation Table"
                  value="simulation"
                ></v-list-item>
                <v-list-item
                  prepend-icon="mdi-table"
                  title="System analysis Table"
                  value="system-analysis"
                  @click="tableSelected = 'system-analysis'"
                ></v-list-item>
              </v-list>
            </v-navigation-drawer>
            <v-main style="height: 250px">
              <div class="intermediate-tab" v-if="level === 'Beginner'">
                <div v-if="tableSelected == 'data-table'">
                  <v-data-table
                    :headers="DataTable['header']"
                    :items="DataTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'arrival-probability'">
                  <v-data-table
                    :headers="arrivalProbabilityTable['header']"
                    :items="arrivalProbabilityTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'service-probability'">
                  <v-data-table
                    :headers="ServiceProbabilityTable['header']"
                    :items="ServiceProbabilityTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'simulation'">
                  <v-data-table
                    :headers="simulationTable['header']"
                    :items="simulationTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'system-analysis'">
                  <v-data-table
                    :headers="SystemAnalysisTable['header']"
                    :items="SystemAnalysisTable.body"
                  >
                  </v-data-table>
                </div>
              </div>
              <div class="intermediate-tab" v-if="level === 'Intermediate'">
                <div v-if="tableSelected == 'data-table'">
                  <v-data-table
                    :headers="DataTable['header']"
                    :items="DataTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'arrival-probability'">
                  <v-data-table
                    :headers="arrivalProbabilityTable['header']"
                    :items="arrivalProbabilityTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'service-probability'">
                  <v-data-table
                    :headers="ServiceProbabilityTable['header']"
                    :items="ServiceProbabilityTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'simulation'">
                  <v-data-table
                    :headers="simulationTable['header']"
                    :items="simulationTable.body"
                  >
                  </v-data-table>
                </div>
                <div v-if="tableSelected == 'system-analysis'">
                  <v-data-table
                    :headers="SystemAnalysisTable['header']"
                    :items="SystemAnalysisTable.body"
                  >
                  </v-data-table>
                </div>
              </div>
            </v-main>
          </v-layout>
        </div>
      </main>
    </v-container>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import { useTheme } from 'vuetify'
import { ref } from 'vue'
import ExcelJS from 'exceljs'
import { useRoute } from 'vue-router'
let level = useRoute().params.level

const theme = useTheme()
const isDarkMode = computed(() => theme.global.name.value === 'dark')
const toggleTheme = () => {
  theme.global.name.value = isDarkMode.value ? 'light' : 'dark'
}
let tableSelected = ref(null)
// Ref to store the Excel data
const columnsData = {}
const simulationTable = ref({
  header: [],
  body: [],
})
const arrivalProbabilityTable = ref({
  header: [],
  body: [],
})
const ServiceProbabilityTable = ref({
  header: [],
  body: [],
})
const DataTable = ref({
  header: [],
  body: [],
})
const SystemAnalysisTable = ref({
  header: [],
  body: [],
})
const formatToClockTime = value => {
  const date = new Date(value)
  if (!isNaN(date.getTime())) {
    // Check if the value is a valid date
    const hours = String(date.getHours()).padStart(2, '0')
    const minutes = String(date.getMinutes()).padStart(2, '0')
    return `${hours}:${minutes}`
  }
  return value // Return the original value if it's not a date
}
// Function to handle file upload
const handleFileUpload = async event => {
  const file = event.target.files[0]
  if (file) {
    const workbook = new ExcelJS.Workbook()
    const arrayBuffer = await file.arrayBuffer()
    await workbook.xlsx.load(arrayBuffer)
    const worksheet = workbook.worksheets[0] // Access the first sheet
    worksheet.columns.forEach(column => {
      const columnHeader = column.values[3] // First row header
      columnsData[columnHeader] = []
      // Start from the second row to skip the header
      column.eachCell((cell, rowNumber) => {
        if (rowNumber > 3) {
          // Skip header row
          columnsData[columnHeader].push(cell.value)
        }
      })
    })
    const columnsKeys = Object.keys(columnsData)
    let colsData = {}
    let idx = 0
    // get headers of table
    for (const key of columnsKeys) {
      colsData[key] = columnsData[key]
      // console.log(colsData[key])
      if (idx <= 10) {
        simulationTable.value['header'].push({ title: key, value: key })
      } else if (idx > 11 && idx <= 16) {
        arrivalProbabilityTable.value['header'].push({ title: key, value: key })
        // console.log('arrivalProbabilityTable', colsData[key])
      } else if (idx >= 17 && idx < 22) {
        ServiceProbabilityTable.value['header'].push({ title: key, value: key })
      } else if (idx >= 22 && idx < 28) {
        DataTable.value['header'].push({ title: key, value: key })
      } else if (idx > 27) {
        SystemAnalysisTable.value['header'].push({ title: key, value: key })
      }
      idx++
    }
    console.log(columnsKeys)
    const headersSimulation = simulationTable.value.header.map(
      header => header.value,
    )
    const headersArrival = arrivalProbabilityTable.value.header.map(
      header => header.value,
    )
    const headersService = ServiceProbabilityTable.value.header.map(
      header => header.value,
    )
    const headersDataTable = DataTable.value.header.map(header => header.value)
    const headersSystemAnalysisTable = SystemAnalysisTable.value.header.map(
      header => header.value,
    )
    // get body of table and convert to array
    for (let i = 0; i < colsData['Customer No'].length; i++) {
      const row = {}
      headersSimulation.forEach(header => {
        let val = colsData[header][i]
        row[header] =
          (val || val == 0) && typeof val !== 'object'
            ? val
            : val?.result != null
              ? val?.result
              : 0
      })
      simulationTable.value.body.push(row)
    }
    for (let i = 0; i < colsData['Time Between Arrivals'].length; i++) {
      const row = {}
      headersArrival.forEach(header => {
        let val = colsData[header][i]
        row[header] =
          (val || val == 0) && typeof val !== 'object'
            ? val
            : val?.result != null
              ? val?.result
              : 0
      })
      arrivalProbabilityTable.value.body.push(row)
    }
    for (let i = 0; i < colsData['Service Time'].length; i++) {
      const row = {}
      headersService.forEach(header => {
        let val = colsData[header][i]
        row[header] =
          (val || val == 0) && typeof val !== 'object'
            ? val
            : val?.result != null
              ? val?.result
              : 0
      })
      ServiceProbabilityTable.value.body.push(row)
    }
    for (let i = 0; i < colsData['ID'].length; i++) {
      const row = {}
      headersDataTable.forEach(header => {
        let val = colsData[header][i]
        if (header == 'arrival time') val = formatToClockTime(val)
        row[header] =
          (val || val == 0) && typeof val !== 'object'
            ? val
            : val?.result != null
              ? val?.result
              : 0
      })
      DataTable.value.body.push(row)
    }
    for (let i = 0; i < colsData['service'].length; i++) {
      const row = {}
      headersSystemAnalysisTable.forEach(header => {
        let val = colsData[header][i]
        row[header] =
          (val || val == 0) && typeof val !== 'object'
            ? val
            : val?.result != null
              ? val?.result
              : 0
      })
      SystemAnalysisTable.value.body.push(row)
    }
  }
}
</script>

<style lang="scss">
.view-analysis {
  padding-top: 50px;
  .dark-light-mode-btn {
    position: fixed;
    right: 20px;
    top: 20px;
    z-index: 9900;
  }
  .go-back {
    position: fixed;
    left: 20px;
    top: 20px;
    z-index: 9900;
  }
  .head {
    position: relative;
    margin-bottom: 30px;
  }
  main {
    .import-file {
    }
  }
  .v-table.v-data-table {
    margin-top: 20px;
    @media (max-width: 992px) {
      margin-top: 100px;
    }
  }
  .v-navigation-drawer__content {
    overflow-y: hidden;
  }
  .table {
    .v-navigation-drawer {
      margin: 0;
      position: absolute !important;
      top: 20px !important;
      transform: translateY(0) !important;
      height: fit-content !important;
      width: 60px !important;
      margin-left: -60px;
      transition: 0.3s;
      &:hover {
        width: 256px !important;
      }
      h3 {
        text-wrap: nowrap;
      }
      @media (max-width: 992px) {
        margin-left: -10px;
        height: 90px !important;
        width: 256px !important;
        top: 5px !important;
        left: 50% !important;
        transform: translateX(-50%) !important;
        // overflow-y: hidden;
        &:hover {
          height: 280px !important;
        }
      }
    }
    .v-data-table-footer {
      display: none;
    }
  }
}
</style>
