<template>
  <div class="hero">
    <v-btn id="go-back" class="go-back" @click="$router.push('/')"
      ><v-icon>mdi-chevron-left</v-icon></v-btn
    >
    <v-container>
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

async function generateSimulationExcel() {
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
  const serviceData = []

  // Read data from the input file
  serviceSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      serviceData.push({
        ID: row.getCell(1).value,
        ArrivalTime: row.getCell(2).value,
        CustName: row.getCell(3).value,
        Service: row.getCell(4).value,
        Duration: row.getCell(5).value,
      })
    }
  })

  // Create a new workbook and sheet for Simulation and Analysis
  const workbook = new ExcelJS.Workbook()
  const simulationSheet = workbook.addWorksheet('Simulation')

  // Write input file data into the simulation sheet
  simulationSheet.getCell('AB1').value = 'Data Table'
  simulationSheet.getCell('Y3').value = 'ID'
  simulationSheet.getCell('Z3').value = 'Arrival Time'
  simulationSheet.getCell('AA3').value = 'Cust Name'
  simulationSheet.getCell('AB3').value = 'Service'
  simulationSheet.getCell('AC3').value = 'Duration'

  serviceData.forEach((row, index) => {
    const rowIndex = index + 4 // Start from row 4 for data entries
    simulationSheet.getCell(`Y${rowIndex}`).value = row.ID || ''
    simulationSheet.getCell(`Z${rowIndex}`).value = row.ArrivalTime || ''
    simulationSheet.getCell(`Z${rowIndex}`).numFmt = 'hh:mm:ss'
    simulationSheet.getCell(`AA${rowIndex}`).value = row.CustName || ''
    simulationSheet.getCell(`AB${rowIndex}`).value = row.Service || ''
    simulationSheet.getCell(`AC${rowIndex}`).value = row.Duration || 0
  })

  // Calculate inter-arrival times and write to Interval column
  simulationSheet.getCell('AD3').value = 'Inter_arrival'
  simulationSheet.getCell('AD4').value = 0 // Initial inter-arrival time

  // Write Arrival Probability table headers
  simulationSheet.getCell('M1').value = 'Arrival Probability'
  simulationSheet.getCell('M3').value = 'Time Between Arrivals'
  simulationSheet.getCell('N3').value = 'Arrival_Probability'
  simulationSheet.getCell('O3').value = 'Arrival_Cumulative '
  simulationSheet.getCell('P2').value = 'Random Digit Assignment'
  simulationSheet.getCell('P3').value = 'Arrival_From'
  simulationSheet.getCell('Q3').value = 'Arrival_To'

  // Function to convert a time format to minutes
  function convertToMinutes(time) {
    if (time instanceof Date) {
      return time.getHours() * 60 + time.getMinutes() // Convert Date object to minutes
    } else if (typeof time === 'string') {
      const parts = time.split(':')
      return parseInt(parts[0]) * 60 + parseInt(parts[1]) // For 'HH:MM' string format
    }
    return 0 // Default return for unexpected formats
  }

  const interArrivalTimes = [] // Store inter-arrival times for later use
  for (let i = 5; i <= serviceData.length + 3; i++) {
    const arrivalTimeCurrent = convertToMinutes(serviceData[i - 4].ArrivalTime)
    const arrivalTimePrevious = serviceData[i - 5]
      ? convertToMinutes(serviceData[i - 5].ArrivalTime)
      : 0
    const interArrival = arrivalTimeCurrent - arrivalTimePrevious
    simulationSheet.getCell(`AD${i}`).value = interArrival // Store in the Inter_arrival column
    interArrivalTimes.push(interArrival) // Collect inter-arrival times (including zero values)
  }

  // Get unique inter-arrival times and sort them
  const uniqueInterArrival = [...new Set(interArrivalTimes)].sort(
    (a, b) => a - b,
  )
  // Write unique inter-arrival times to the Arrival Probability table
  uniqueInterArrival.forEach((value, index) => {
    const rowIndex = index + 4 // Start writing from row 4 for unique values
    simulationSheet.getCell(`M${rowIndex}`).value = value // Time Between Arrivals
    simulationSheet.getCell(`N${rowIndex}`).value = {
      formula: `COUNTIF($AD$4:$AD$${serviceData.length + 3}, M${rowIndex}) / COUNT($AD$4:$AD$${serviceData.length + 3})`,
    } // Probability
    if (rowIndex != 4)
      simulationSheet.getCell(`O${rowIndex}`).value = {
        formula: `O${rowIndex - 1} + N${rowIndex}`,
      } // Cumulative Probability
    else simulationSheet.getCell(`O${rowIndex}`).value = { formula: `N4` } // Cumulative Probability
    if (rowIndex != 4)
      simulationSheet.getCell(`P${rowIndex}`).value = {
        formula: `Q${rowIndex - 1}+1`,
      }
    else simulationSheet.getCell(`P${rowIndex}`).value = 1
    simulationSheet.getCell(`Q${rowIndex}`).value = {
      formula: `O${rowIndex}*100`,
    }
  })

  // Write Service Probability table
  simulationSheet.getCell('U1').value = 'Service Probability'
  simulationSheet.getCell('V3').value = 'Service_From'
  simulationSheet.getCell('W3').value = 'Service_To'
  simulationSheet.getCell('S3').value = 'Service Time'
  simulationSheet.getCell('T3').value = 'Service_Probability'
  simulationSheet.getCell('U3').value = 'Service_Cumulative '
  simulationSheet.getCell('V2').value = 'Random Digit Assignment'
  const uniqueDurations = [
    ...new Set(serviceData.map(row => row.Duration)),
  ].sort((a, b) => a - b)
  uniqueDurations.forEach((value, index) => {
    const rowIndex = index + 4 // Start from row 4 for unique values
    simulationSheet.getCell(`S${rowIndex}`).value = value
    simulationSheet.getCell(`T${rowIndex}`).value = {
      formula: `COUNTIF($AC$4:$AC$${serviceData.length + 3}, S${rowIndex}) / COUNT($AC$4:$AC$${serviceData.length + 3})`,
    }
    if (rowIndex != 4)
      simulationSheet.getCell(`U${rowIndex}`).value = {
        formula: `U${rowIndex - 1} + T${rowIndex}`,
      }
    else
      simulationSheet.getCell(`U${rowIndex}`).value = {
        formula: `T${rowIndex}`,
      }
    if (rowIndex == 4)
      simulationSheet.getCell(`V${rowIndex}`).value = 1 // Start of random digit assignment
    else
      simulationSheet.getCell(`V${rowIndex}`).value = {
        formula: `W${rowIndex - 1}+1`,
      }
    simulationSheet.getCell(`W${rowIndex}`).value = {
      formula: `U${rowIndex} * 100`,
    }
  })

  // Write Simulation Table
  simulationSheet.getCell('E1').value = 'Simulation Table'
  simulationSheet.getCell('A3').value = 'Customer No'
  simulationSheet.getCell('B3').value = 'Rand Digit Interval'
  simulationSheet.getCell('C3').value = 'Interval Time'
  simulationSheet.getCell('D3').value = 'Arr Clock'
  simulationSheet.getCell('E3').value = 'Start Time'
  simulationSheet.getCell('F3').value = 'Rand Digit Service'
  simulationSheet.getCell('G3').value = 'Ser Duration'
  simulationSheet.getCell('H3').value = 'End Time'
  simulationSheet.getCell('I3').value = 'Cust Waiting Time'
  simulationSheet.getCell('J3').value = 'Server Idle Time'
  simulationSheet.getCell('K3').value = 'Cust Spent Time'
  simulationSheet.getCell('A4').value = 1
  simulationSheet.getCell('B4').value = { formula: `RANDBETWEEN(1, 100)` }
  simulationSheet.getCell('C4').value = 0
  simulationSheet.getCell('D4').value = 0
  simulationSheet.getCell('E4').value = 0
  simulationSheet.getCell('F4').value = { formula: `RANDBETWEEN(1, 100)` }
  simulationSheet.getCell('G4').value = {
    formula: `LOOKUP(F4, $V$4:$W${uniqueDurations.length + 3}, $S$4:$S${uniqueDurations.length + 3})`,
  }
  simulationSheet.getCell('H4').value = { formula: `E4+G4` }
  simulationSheet.getCell('I4').value = { formula: `E4-D4` }
  simulationSheet.getCell('J4').value = { formula: `E5-H4` }
  simulationSheet.getCell('K4').value = { formula: `H4-D4` }

  for (let i = 5; i <= noCustomers.value; i++) {
    simulationSheet.getCell(`A${i}`).value = { formula: `A${i - 1} + 1` }
    simulationSheet.getCell(`B${i}`).value = { formula: `RANDBETWEEN(1, 100)` }
    simulationSheet.getCell(`C${i}`).value = {
      formula: `LOOKUP(B${i},$P$4:$Q$8,$M$4:$M$8)`,
    }
    simulationSheet.getCell(`D${i}`).value = { formula: `D${i - 1} + C${i}` }
    simulationSheet.getCell(`E${i}`).value = {
      formula: `MAX(H${i - 1}, D${i})`,
    }
    simulationSheet.getCell(`F${i}`).value = { formula: `RANDBETWEEN(1, 100)` }
    simulationSheet.getCell(`G${i}`).value = {
      formula: `LOOKUP(F${i}, $V$4:$W${uniqueDurations.length + 3}, $S$4:$S${uniqueDurations.length + 3})`,
    }
    simulationSheet.getCell(`H${i}`).value = { formula: `E${i} + G${i}` }
    simulationSheet.getCell(`I${i}`).value = { formula: `E${i} - D${i}` }
    simulationSheet.getCell(`J${i}`).value = { formula: `E${5}-H${4}` }
    simulationSheet.getCell(`K${i}`).value = { formula: `H${i} - D${i}` }
  }

  // Write System Analysis Table
  simulationSheet.getCell('AG1').value = 'System Analysis Table'
  simulationSheet
  simulationSheet.getCell('AF3').value = 'Metric'
  simulationSheet.getCell('AG3').value = 'Value'
  simulationSheet.getCell('AF4').value = 'Number of Customers'
  simulationSheet.getCell('AG4').value = {
    formula: `COUNT(A4:A${noCustomers.value + 3})`,
  }
  simulationSheet.getCell('AF5').value = 'Total Waiting'
  simulationSheet.getCell('AG5').value = {
    formula: `SUM(I4:I${noCustomers.value + 3})`,
  }
  simulationSheet.getCell('AF6').value = 'Avg Waiting'
  simulationSheet.getCell('AG6').value = { formula: `AG5/AG4` }
  simulationSheet.getCell('AF7').value = 'Probability of Waiting'
  simulationSheet.getCell('AG7').value = {
    formula: `COUNTIF(I4:I${noCustomers.value + 3}, ">0") / AG4`,
  }
  simulationSheet.getCell('AF8').value = 'Number of Idle'
  simulationSheet.getCell('AG8').value = {
    formula: `COUNTIF(J4:J${noCustomers.value + 3}, ">0")`,
  }
  simulationSheet.getCell('AF9').value = 'Number of Busy'
  simulationSheet.getCell('AG9').value = {
    formula: `COUNTIF(J4:J${noCustomers.value + 3}, "0")`,
  }
  simulationSheet.getCell('AF10').value = 'Probability of Idle'
  simulationSheet.getCell('AG10').value = { formula: `AG8/AG4` }
  simulationSheet.getCell('AF11').value = 'Probability of Busy'
  simulationSheet.getCell('AG11').value = { formula: `AG9/AG4` }
  simulationSheet.getCell('AF12').value = 'Avg Ser Duration'
  simulationSheet.getCell('AG12').value = {
    formula: `SUM(G4:G${noCustomers.value + 3})/AG4`,
  }
  simulationSheet.getCell('AF13').value = 'Avg Interval'
  simulationSheet.getCell('AG13').value = {
    formula: `SUM(C4:C${noCustomers.value + 3})/(AG4-1)`,
  }

  // Write the output Excel file
  const excelBuffer = await workbook.xlsx.writeBuffer() // Use writeBuffer
  const blob = new Blob([excelBuffer], { type: 'application/octet-stream' })
  saveAs(blob, 'Simulation.xlsx') // Prompt download
}

async function generateSimulationExcel2() {
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
  const serviceData = []

  // Read data from the input file
  serviceSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      serviceData.push({
        ID: row.getCell(1).value,
        ArrivalTime: row.getCell(2).value,
        CustName: row.getCell(3).value,
        Service: row.getCell(4).value,
        Duration: row.getCell(5).value,
        Server: row.getCell(6).value,
      })
    }
  })

  // Create a new workbook and sheet for Simulation and Analysis
  const workbook = new ExcelJS.Workbook()
  const simulationSheet = workbook.addWorksheet('Simulation')

  // Write input file data into the simulation sheet
  simulationSheet.getCell('AL1').value = 'Data Table'
  simulationSheet.getCell('AI3').value = 'ID'
  simulationSheet.getCell('AJ3').value = 'Arrival Time'
  simulationSheet.getCell('AK3').value = 'Cust Name'
  simulationSheet.getCell('AL3').value = 'Service'
  simulationSheet.getCell('AM3').value = 'Duration'
  simulationSheet.getCell('AN3').value = 'Server'

  serviceData.forEach((row, index) => {
    const rowIndex = index + 4 // Start from row 4 for data entries
    simulationSheet.getCell(`AI${rowIndex}`).value = row.ID || ''
    simulationSheet.getCell(`AJ${rowIndex}`).value = row.ArrivalTime || ''
    simulationSheet.getCell(`AJ${rowIndex}`).numFmt = 'hh:mm:ss'
    simulationSheet.getCell(`AK${rowIndex}`).value = row.CustName || ''
    simulationSheet.getCell(`AL${rowIndex}`).value = row.Service || ''
    simulationSheet.getCell(`AM${rowIndex}`).value = row.Duration || 0
    simulationSheet.getCell(`AN${rowIndex}`).value = row.Server || ''
  })

  // Calculate inter-arrival times and write to Interval column
  simulationSheet.getCell('AO3').value = 'Inter_arrival'
  simulationSheet.getCell('AO4').value = 0

  // Write Arrival Probability table headers
  simulationSheet.getCell('R1').value = 'Arrival Probability'
  simulationSheet.getCell('P3').value = 'Time Between Arrivals'
  simulationSheet.getCell('Q3').value = 'Probability'
  simulationSheet.getCell('R3').value = 'Cumulative'
  simulationSheet.getCell('S3').value = 'Random Digit Assignment'
  simulationSheet.getCell('S2').value = 'From'
  simulationSheet.getCell('T2').value = 'To'

  // Function to convert a time format to minutes
  function convertToMinutes(time) {
    if (time instanceof Date) {
      return time.getHours() * 60 + time.getMinutes() // Convert Date object to minutes
    } else if (typeof time === 'string') {
      const parts = time.split(':')
      return parseInt(parts[0]) * 60 + parseInt(parts[1]) // For 'HH:MM' string format
    }
    return 0 // Default return for unexpected formats
  }

  const interArrivalTimes = [] // Store inter-arrival times for later use
  for (let i = 5; i <= serviceData.length + 3; i++) {
    const arrivalTimeCurrent = convertToMinutes(serviceData[i - 4].ArrivalTime)
    const arrivalTimePrevious = serviceData[i - 5]
      ? convertToMinutes(serviceData[i - 5].ArrivalTime)
      : 0
    const interArrival = arrivalTimeCurrent - arrivalTimePrevious
    simulationSheet.getCell(`AO${i}`).value = interArrival // Store in the Inter_arrival column
    interArrivalTimes.push(interArrival) // Collect inter-arrival times (including zero values)
  }

  // Get unique inter-arrival times and sort them
  const uniqueInterArrival = [...new Set(interArrivalTimes)].sort(
    (a, b) => a - b,
  )
  // Write unique inter-arrival times to the Arrival Probability table
  uniqueInterArrival.forEach((value, index) => {
    const rowIndex = index + 4 // Start writing from row 4 for unique values
    simulationSheet.getCell(`P${rowIndex}`).value = value // Time Between Arrivals
    simulationSheet.getCell(`Q${rowIndex}`).value = {
      formula: `COUNTIF($AO$4:$AO$${serviceData.length + 3}, P${rowIndex}) / COUNT($AO$4:$AO$${serviceData.length + 3})`,
    } // Probability
    if (rowIndex != 4)
      simulationSheet.getCell(`R${rowIndex}`).value = {
        formula: `R${rowIndex - 1} + Q${rowIndex}`,
      } // Cumulative Probability
    else simulationSheet.getCell(`R${rowIndex}`).value = { formula: `Q4` } // Cumulative Probability
    if (rowIndex != 4)
      simulationSheet.getCell(`S${rowIndex}`).value = {
        formula: `T${rowIndex - 1}+1`,
      }
    else simulationSheet.getCell(`S${rowIndex}`).value = 1
    simulationSheet.getCell(`T${rowIndex}`).value = {
      formula: `R${rowIndex}*100`,
    }
  })

  // Write server 1 Probability table
  simulationSheet.getCell('X1').value = 'Server_01'
  simulationSheet.getCell('Y2').value = 'From'
  simulationSheet.getCell('Z2').value = 'To'
  simulationSheet.getCell('V3').value = 'Service Time'
  simulationSheet.getCell('W3').value = 'Probability'
  simulationSheet.getCell('X3').value = 'Cumulative'
  simulationSheet.getCell('Y3').value = 'Random Digit Assignment'
  const uniqueDurations1 = [
    ...new Set(
      serviceData
        .filter(row => row.Server === 'ser01') // filter by server type
        .map(row => row.Duration),
    ),
  ].sort((a, b) => a - b)
  uniqueDurations1.forEach((value, index) => {
    const rowIndex = index + 4 // Start from row 4 for unique values
    simulationSheet.getCell(`V${rowIndex}`).value = value
    simulationSheet.getCell(`W${rowIndex}`).value = {
      formula: `COUNTIFS($AM$${4}:$AM$${serviceData.length + 3}, V${rowIndex}, $AN$${4}:$AN$${serviceData.length + 3}, "ser01") / COUNTIF($AN$${4}:$AN$${serviceData.length + 3}, "ser01")`,
    }

    if (rowIndex != 4)
      simulationSheet.getCell(`X${rowIndex}`).value = {
        formula: `X${rowIndex - 1} + W${rowIndex}`,
      }
    else
      simulationSheet.getCell(`X${rowIndex}`).value = {
        formula: `W${rowIndex}`,
      }
    if (rowIndex == 4)
      simulationSheet.getCell(`Y${rowIndex}`).value = 1 // Start of random digit assignment
    else
      simulationSheet.getCell(`Y${rowIndex}`).value = {
        formula: `Z${rowIndex - 1}+1`,
      }
    simulationSheet.getCell(`Z${rowIndex}`).value = {
      formula: `X${rowIndex} * 100`,
    }
  })

  // Write server 2 Probability table
  simulationSheet.getCell('AD1').value = 'Server_02'
  simulationSheet.getCell('AE2').value = 'From'
  simulationSheet.getCell('AF2').value = 'To'
  simulationSheet.getCell('AB3').value = 'Service Time'
  simulationSheet.getCell('AC3').value = 'Probability'
  simulationSheet.getCell('AD3').value = 'Cumulative'
  simulationSheet.getCell('AE3').value = 'Random Digit Assignment'

  const uniqueDurations = [
    ...new Set(
      serviceData
        .filter(row => row.Server === 'ser02') // filter by server type
        .map(row => row.Duration),
    ),
  ].sort((a, b) => a - b)

  uniqueDurations.forEach((value, index) => {
    const rowIndex = index + 4 // Start from row 4 for unique values
    simulationSheet.getCell(`AB${rowIndex}`).value = value
    simulationSheet.getCell(`AC${rowIndex}`).value = {
      formula: `COUNTIFS($AM$${4}:$AM$${serviceData.length + 3}, AB${rowIndex}, $AN$${4}:$AN$${serviceData.length + 3}, "ser02") / COUNTIF($AN$${4}:$AN$${serviceData.length + 3}, "ser02")`,
    }
    if (rowIndex != 4)
      simulationSheet.getCell(`AD${rowIndex}`).value = {
        formula: `AD${rowIndex - 1} + AC${rowIndex}`,
      }
    else
      simulationSheet.getCell(`AD${rowIndex}`).value = {
        formula: `AC${rowIndex}`,
      }
    if (rowIndex == 4)
      simulationSheet.getCell(`AE${rowIndex}`).value = 1 // Start of random digit assignment
    else
      simulationSheet.getCell(`AE${rowIndex}`).value = {
        formula: `AF${rowIndex - 1}+1`,
      }
    simulationSheet.getCell(`AF${rowIndex}`).value = {
      formula: `AD${rowIndex} * 100`,
    }
  })

  // Write Simulation Table
  simulationSheet.getCell('H1').value = 'Simulation Table'
  simulationSheet.getCell('A3').value = 'Customer No'
  simulationSheet.getCell('B3').value = 'Rand Digit Interval'
  simulationSheet.getCell('C3').value = 'Interval Time'
  simulationSheet.getCell('D3').value = 'Arr Clock'
  simulationSheet.getCell('E3').value = 'Rand Digit Service'
  simulationSheet.getCell('F3').value = 'Start'
  simulationSheet.getCell('G3').value = 'Duration'
  simulationSheet.getCell('H3').value = 'End'
  simulationSheet.getCell('I3').value = 'Start'
  simulationSheet.getCell('J3').value = 'Duration'
  simulationSheet.getCell('K3').value = 'End'
  simulationSheet.getCell('L3').value = 'Cust Waiting Time'
  simulationSheet.getCell('M3').value = 'Ser idle'

  for (let i = 5; i <= noCustomers.value; i++) {
    if (i == 5) {
      simulationSheet.getCell(`A${i - 1}`).value = 1
      simulationSheet.getCell(`C${i - 1}`).value = 0
      simulationSheet.getCell(`D${i - 1}`).value = 0
      simulationSheet.getCell(`E${i - 1}`).value = {
        formula: `RANDBETWEEN(1, 100)`,
      }
      simulationSheet.getCell(`F${i - 1}`).value = 0
      simulationSheet.getCell(`G${i - 1}`).value = {
        formula: `LOOKUP(E4,Y4:Z${uniqueDurations1.length + 3},V4:V${uniqueDurations1.length + 3})`,
      }
      simulationSheet.getCell(`H${i - 1}`).value = { formula: `F4+G4` }
      simulationSheet.getCell(`L${i - 1}`).value = {
        formula: `MINUTE(IF(F4<>"",I4-D4,F4-D4))`,
      }
      simulationSheet.getCell(`M${i - 1}`).value = {
        formula: `IF(F4<>"","Server_01","Server_02")`,
      }
    }

    simulationSheet.getCell(`A${i}`).value = { formula: `A${i - 1}+1` }
    simulationSheet.getCell(`B${i}`).value = { formula: `RANDBETWEEN(1, 100)` }
    simulationSheet.getCell(`C${i}`).value = {
      formula: `LOOKUP(B${i},$S$4:$T${uniqueDurations.length + 3},$P$4:$P$${uniqueDurations.length + 3})`,
    }
    simulationSheet.getCell(`D${i}`).value = { formula: `D${i - 1}+C${i}` }
    simulationSheet.getCell(`E${i}`).value = { formula: `RANDBETWEEN(1, 100)` }
    simulationSheet.getCell(`F${i}`).value = {
      formula: `IF(MAX('$H$4':H${i - 1})>MAX('$K$4:K${i - 1}),"",MAX($H$4:H$4,D${i}))`,
    }
    simulationSheet.getCell(`G${i}`).value = {
      formula: `LOOKUP(E${i},Y4:Z${uniqueDurations.length + 3},V4:V${uniqueDurations.length + 3})`,
    }
    simulationSheet.getCell(`H${i}`).value = {
      formula: `IF(F${i}<>"",F${i}+G${i},"")`,
    }
    simulationSheet.getCell(`I${i}`).value = {
      formula: `IF(F${i}<>"","",MAX($K$${i - 1}:K${i - 1},D${i}))`,
    }
    simulationSheet.getCell(`J${i}`).value = {
      formula: `IF(I${i}<>"",LOOKUP(E${i},$AE$4:$AF$${uniqueDurations.length + 3},$AB$4:$AB$${uniqueDurations.length + 3}),"")`,
    }

    simulationSheet.getCell(`K${i}`).value = {
      formula: `IF(I${i}<>"",I${i}+J${i},"")`,
    }
    simulationSheet.getCell(`L${i}`).value = {
      formula: `MINUTE(IF(F${i - 1}<>"",F${i - 1}-D${i - 1},I${i - 1}-D${i - 1}))`,
    }
    simulationSheet.getCell(`M${i}`).value = {
      formula: `IF(F${i - 1}<>"","Server_01","Server_02")`,
    }
  }

  // Write the output Excel file
  const excelBuffer = await workbook.xlsx.writeBuffer() // Use writeBuffer
  const blob = new Blob([excelBuffer], { type: 'application/octet-stream' })
  saveAs(blob, 'Simulation.xlsx') // Prompt download
}

const SimulateData = async () => {
  if (level === 'Beginner') {
    await createSimulationExcel() // Call the function when level is Beginner
  } else if (level == 'Intermediate') {
    await generateSimulationExcel()
  } else {
    await generateSimulationExcel2()
  }
  // You can add more logic for other levels here
}

const onFileChange = event => {
  excelFile.value = event.target.files[0] // Store the uploaded file
}
</script>
<style lang="scss">
#go-back {
  position: fixed;
  left: 20px;
  top: 20px;
  z-index: 999;
}
.tables {
  position: relative;
}
</style>
