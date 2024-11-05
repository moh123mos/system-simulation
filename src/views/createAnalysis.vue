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
      <v-form @submit.prevent>
        <div class="inputs">
          <div class="no-customers">
            <v-text-field
              label="Number Of Customers"
              type="number"
              v-model="noCustomers"
              :min="1"
            ></v-text-field>
          </div>
          <div v-if="level === 'Beginner'" class="min-inter-arrival">
            <v-text-field
              label="Min Interarrival "
              type="number"
              v-model="minInterarrival"
              :min="1"
            ></v-text-field>
          </div>
          <div v-if="level === 'Beginner'" class="max-inter-arrival">
            <v-text-field
              label="Max Interarrival"
              type="number"
              v-model="maxInterarrival"
              :min="1"
            ></v-text-field>
          </div>
          <div class="upload-service">
            <input
              @change="onFileChange"
              id="upload-file"
              type="file"
              title="Upload Service Table"
            />
          </div>
        </div>
        <div class="text-center">
          <v-btn
            type="submit"
            @click="SimulateData"
            style="text-transform: capitalize"
          >
            Simulate</v-btn
          >
        </div>
      </v-form>
    </v-container>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'
import { useRoute } from 'vue-router'
import { useTheme } from 'vuetify'
import ExcelJS from 'exceljs' // Import ExcelJS
import { saveAs } from 'file-saver' // Ensure this is at the top of your script

//import fs from 'fs'; // Node.js fs module

const theme = useTheme()
const isDarkMode = computed(() => theme.global.name.value === 'dark')
const toggleTheme = () => {
  theme.global.name.value = isDarkMode.value ? 'light' : 'dark'
}

const route = useRoute()
let level = route.params.level
let noCustomers = ref(null)
let minInterarrival = ref(null)
let maxInterarrival = ref(null)
let excelFile = ref(null)

async function createSimulationExcel() {
  // Check if a file has been uploaded
  if (!excelFile.value) {
    alert('Please upload a service table file first.')
    return
  }

  const fileReader = new FileReader()

  // Read the uploaded file as an ArrayBuffer
  const fileContent = await new Promise((resolve, reject) => {
    fileReader.onload = event => resolve(event.target.result)
    fileReader.onerror = error => reject(error)
    fileReader.readAsArrayBuffer(excelFile.value)
  })

  // Initialize a new ExcelJS workbook
  const serviceWorkbook = new ExcelJS.Workbook()

  // Load the input Excel file into the workbook
  try {
    await serviceWorkbook.xlsx.load(fileContent)
  } catch (error) {
    console.error('Error loading Excel file:', error)
    alert('Failed to load the Excel file. Please check the format.')
    return
  }

  // Assume the service table is in the first sheet of the input file
  const serviceSheet = serviceWorkbook.worksheets[0]
  const serviceTable = []

  // Read data from Service Table in input file
  serviceSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      serviceTable.push({
        SerID: row.getCell(1).value,
        Service: row.getCell(2).value,
        SerDuration: row.getCell(3).value,
      })
    }
  })

  // Debug: Log the service table to see what’s being read
  console.log('Service Table:', serviceTable)

  // Create a new workbook and sheet for Simulation and Analysis
  const workbook = new ExcelJS.Workbook()
  const simulationSheet = workbook.addWorksheet('Simulation')

  // Set up headers for the service table
  simulationSheet.getCell('O3').value = 'Service Table'
  simulationSheet.getCell('N5').value = 'Ser.ID'
  simulationSheet.getCell('O5').value = 'Service'
  simulationSheet.getCell('P5').value = 'Ser.Duration'

  // Write the entire service table columns to the simulation sheet
  serviceTable.forEach((row, index) => {
    const rowIndex = index + 6 // Start writing from row 6 for the service table

    // Write each column's value
    simulationSheet.getCell(`N${rowIndex}`).value = row.SerID || '' // Service ID
    simulationSheet.getCell(`O${rowIndex}`).value = row.Service || '' // Service Name
    simulationSheet.getCell(`P${rowIndex}`).value = row.SerDuration || 0 // Service Duration
  })

  // Set up Simulation Table headers
  simulationSheet.getCell('G3').value = 'Simulation Table'
  simulationSheet.getCell('A5').value = 'Customer'
  simulationSheet.getCell('B5').value = 'Interarrival'
  simulationSheet.getCell('C5').value = 'Arrival Time'
  simulationSheet.getCell('D5').value = 'Service Code'
  simulationSheet.getCell('E5').value = 'Service Title'
  simulationSheet.getCell('F5').value = 'Service Begins'
  simulationSheet.getCell('G5').value = 'Service Duration'
  simulationSheet.getCell('H5').value = 'Service End'
  simulationSheet.getCell('I5').value = 'System State'
  simulationSheet.getCell('J5').value = 'Customer State'
  simulationSheet.getCell('K5').value = 'Waiting Time'

  // Fill in simulation data
  for (let i = 0; i < noCustomers.value; i++) {
    const rowIndex = i + 6

    // Customer Number
    if (i == 0) {
      simulationSheet.getCell(`A${rowIndex}`).value = 1
    } else {
      simulationSheet.getCell(`A${rowIndex}`).value = {
        formula: `A${rowIndex - 1}+1`,
      }
    }

    // Interarrival Time (Random)
    if (i === 0) {
      simulationSheet.getCell(`B${rowIndex}`).value = 0 // First customer has 0 interarrival time
    } else {
      simulationSheet.getCell(`B${rowIndex}`).value = {
        formula: `RANDBETWEEN(${minInterarrival.value}, ${maxInterarrival.value})`,
      }
    }

    // Arrival Time
    if (i === 0) {
      simulationSheet.getCell(`C${rowIndex}`).value = 0
    } else {
      simulationSheet.getCell(`C${rowIndex}`).value = {
        formula: `C${rowIndex - 1} + B${rowIndex}`,
      }
    }

    // Random Service Code
    simulationSheet.getCell(`D${rowIndex}`).value = {
      formula: `RANDBETWEEN(MIN(N6:N${5 + serviceTable.length}), MAX(N6:N${5 + serviceTable.length}))`,
    }

    // Service Title
    simulationSheet.getCell(`E${rowIndex}`).value = {
      formula: `LOOKUP(D${rowIndex}, $N$6:$N$${5 + serviceTable.length}, $O$6:$O$${5 + serviceTable.length})`,
    }

    // Service Begins
    if (i === 0) {
      simulationSheet.getCell(`F${rowIndex}`).value = 0
    } else {
      simulationSheet.getCell(`F${rowIndex}`).value = {
        formula: `MAX(C${rowIndex}, H${rowIndex - 1})`,
      }
    }

    // Service Duration
    simulationSheet.getCell(`G${rowIndex}`).value = {
      formula: `LOOKUP(D${rowIndex}, $N$6:$N$${5 + serviceTable.length}, $P$6:$P$${5 + serviceTable.length})`,
    }

    // Service End
    simulationSheet.getCell(`H${rowIndex}`).value = {
      formula: `F${rowIndex} + G${rowIndex}`,
    }

    // System State
    simulationSheet.getCell(`I${rowIndex}`).value =
      i === 0
        ? 'Busy'
        : { formula: `IF(F${rowIndex} > H${rowIndex - 1}, "Idle", "Busy")` }

    // Customer State
    simulationSheet.getCell(`J${rowIndex}`).value = {
      formula: `IF(C${rowIndex} < F${rowIndex}, "Waiting", "InService")`,
    }

    // Waiting Time
    simulationSheet.getCell(`K${rowIndex}`).value = {
      formula: `F${rowIndex} - C${rowIndex}`,
    }
  }

  // --- System Analysis Table ---
  simulationSheet.getCell('T3').value = 'System Analysis Table'
  simulationSheet.getCell('S5').value = 'Number of Customers'
  simulationSheet.getCell('T5').value = {
    formula: `COUNT(A6:A${5 + noCustomers.value})`,
  }

  simulationSheet.getCell('S6').value = 'Number of Customers Waiting'
  simulationSheet.getCell('T6').value = {
    formula: `COUNTIF(J6:J${5 + noCustomers.value}, "Waiting")`,
  }

  simulationSheet.getCell('S7').value = 'Total Waiting Time'
  simulationSheet.getCell('T7').value = {
    formula: `SUM(K6:K${5 + noCustomers.value})`,
  }

  simulationSheet.getCell('S8').value = 'Average Waiting Time'
  simulationSheet.getCell('T8').value = { formula: `T7/T5` }

  simulationSheet.getCell('S9').value = 'Probability of Waiting'
  simulationSheet.getCell('T9').value = { formula: `T6/T5` }

  simulationSheet.getCell('S10').value = 'Number of Idle States'
  simulationSheet.getCell('T10').value = {
    formula: `COUNTIF(I6:I${5 + noCustomers.value}, "Idle")`,
  }

  simulationSheet.getCell('S11').value = 'Number of Busy States'
  simulationSheet.getCell('T11').value = {
    formula: `COUNTIF(I6:I${5 + noCustomers.value}, "Busy")`,
  }

  simulationSheet.getCell('S12').value = 'Probability of Being Busy'
  simulationSheet.getCell('T12').value = { formula: `T11/T5` }

  simulationSheet.getCell('S13').value = 'Probability of Being Idle'
  simulationSheet.getCell('T13').value = { formula: `T10/T5` }

  // Save the workbook to a Blob
  try {
    const excelBuffer = await workbook.xlsx.writeBuffer() // Use writeBuffer
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' })
    saveAs(blob, 'Simulation.xlsx') // Prompt download
    alert('Simulation results saved successfully.')
  } catch (error) {
    console.error('Error saving the Excel file:', error)
    alert('Failed to save the Excel file. Please try again.')
  }
}

const SimulateData = async () => {
  if (level === 'Beginner') {
    await createSimulationExcel() // Call the function when level is Beginner
  }
  // You can add more logic for other levels here
}

const onFileChange = event => {
  excelFile.value = event.target.files[0] // Store the uploaded file
}
</script>
<style lang="scss">
.go-back {
  position: fixed;
  left: 20px;
  top: 20px;
}
.tables {
  position: relative;
}
</style>
