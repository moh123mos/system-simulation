<template>
  <div>
    <div class="flex-1 p:2 sm:p-6 justify-between flex flex-col">
      <div
        id="messages"
        class="flex flex-col w-[90%] space-y-4 p-3 overflow-y-auto scrollbar-thumb-blue scrollbar-thumb-rounded scrollbar-track-blue-lighter scrollbar-w-2 scrolling-touch pb-[80px]"
      >
        <div
          class="chat-message w-full"
          v-for="(message, i) in messages"
          :key="i"
        >
          <div
            class="flex items-end gap-3"
            :class="!(i % 2) ? 'justify-end' : ''"
          >
            <img
              src="/simulator.jpg"
              alt="My profile"
              class="w-6 h-6 rounded-full"
              :class="!(i % 2) && 'hidden'"
            />
            <div class="flex justify-end">
              <div class="bg-blue-200 text-black p-2 rounded-lg max-w-[600px]">
                {{ message }}
              </div>
            </div>
          </div>
        </div>
      </div>
      <div
        class="fixed bottom-0 w-[80%] left-[50%] translate-x-[-50%] bg-white dark:bg-black text-black dark:text-white duration-300 px-4 py-4 mb-2 sm:mb-0"
      >
        <div class="relative flex">
          <input
            type="text"
            v-model="userMessage"
            @keydown.enter="askOpenAI"
            placeholder="Write your message!"
            class="w-full focus:outline-none focus:placeholder-gray-400 text-gray-600 placeholder-gray-600 pl-12 bg-gray-200 rounded-md py-3"
          />
          <div class="absolute right-0 items-center inset-y-0 hidden sm:flex">
            <label
              type="button"
              class="inline-flex items-center justify-center rounded-full h-10 w-10 transition duration-500 ease-in-out text-gray-500 hover:bg-gray-300 focus:outline-none"
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                class="h-6 w-6 text-gray-600"
              >
                <path
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  stroke-width="2"
                  d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13"
                ></path>
              </svg>
              <input
                @change="handleFileUpload"
                id="dropzone-file"
                type="file"
                title="Upload Service Table"
                accept=".xlsx, .xls"
                class="hidden"
              />
            </label>
            <button
              type="button"
              class="inline-flex items-center justify-center rounded-full h-10 w-10 transition duration-500 ease-in-out text-gray-500 hover:bg-gray-300 focus:outline-none"
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                class="h-6 w-6 text-gray-600"
              >
                <path
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  stroke-width="2"
                  d="M14.828 14.828a4 4 0 01-5.656 0M9 10h.01M15 10h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                ></path>
              </svg>
            </button>
            <button
              type="button"
              @click="askOpenAI"
              class="inline-flex items-center justify-center rounded-lg px-4 py-3 transition duration-500 ease-in-out text-white bg-blue-500 hover:bg-blue-400 focus:outline-none"
            >
              <span class="font-bold">Send</span>
              <svg
                xmlns="http://www.w3.org/2000/svg"
                viewBox="0 0 20 20"
                fill="currentColor"
                class="h-6 w-6 ml-2 transform rotate-90"
              >
                <path
                  d="M10.894 2.553a1 1 0 00-1.788 0l-7 14a1 1 0 001.169 1.409l5-1.429A1 1 0 009 15.571V11a1 1 0 112 0v4.571a1 1 0 00.725.962l5 1.428a1 1 0 001.17-1.408l-7-14z"
                ></path>
              </svg>
            </button>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>
<script setup>
import { ref } from 'vue'
import ExcelJS from 'exceljs'
const userMessage = ref('')
let messages = ref([])
const response = ref('')

const tableData = `
Customer Interarrival Arrival_Time Service_Code Service_Title Service_Begins Service_Duration Service_End System_State Customer_State Waiting_Time
1 0 0 7 Rport 0 4 4 Busy InService 0
2 1 1 1 Open 4 10 14 Busy Waiting 3
3 4 5 4 Withdraw 14 7 21 Busy Waiting 9
4 12 17 5 Transfer 21 8 29 Busy Waiting 4
5 12 29 4 Withdraw 29 7 36 Busy InService 0
6 20 49 2 Delete 49 15 64 Idle InService 0
7 12 61 3 Deposite 64 5 69 Busy Waiting 3
8 18 79 2 Delete 79 15 94 Idle InService 0
9 13 92 7 Rport 94 4 98 Busy Waiting 2
10 16 108 1 Open 108 10 118 Idle InService 0
11 16 124 3 Deposite 124 5 129 Idle InService 0
12 13 137 5 Transfer 137 8 145 Idle InService 0
13 1 138 3 Deposite 145 5 150 Busy Waiting 7
14 7 145 4 Withdraw 150 7 157 Busy Waiting 5
15 18 163 2 Delete 163 15 178 Idle InService 0
16 12 175 6 Inquiry 178 3 181 Busy Waiting 3
17 11 186 5 Transfer 186 8 194 Idle InService 0
18 13 199 4 Withdraw 199 7 206 Idle InService 0
19 5 204 3 Deposite 206 5 211 Busy Waiting 2
20 20 224 1 Open 224 10 234 Idle InService 0
`;

const askOpenAI = async () => {
  try {
    // Combine user message with table data
    const combinedMessage = `
      The following table represents a simulation of a service system. Answer questions based on this data:

      ${tableData}

      User Query: ${userMessage.value}
    `;

    messages.value.push(userMessage.value);

    const res = await fetch('http://localhost:3000/api/chat', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ message: combinedMessage }),
    });

    const data = await res.json();
    response.value = data?.reply;
    messages.value.push(response.value);
  } catch (error) {
    console.error('Error:', error);
    response.value = 'Failed to fetch response.';
  }
};


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

    // const rowsData = parsedRows // Update the reactive state
    // let dataTable = parsedRows.slice(1) // Update the reactive state
  }
}
</script>
<style>
.scrollbar-w-2::-webkit-scrollbar {
  width: 0.25rem;
  height: 0.25rem;
}

.scrollbar-track-blue-lighter::-webkit-scrollbar-track {
  --bg-opacity: 1;
  background-color: #f7fafc;
  background-color: rgba(247, 250, 252, var(--bg-opacity));
}

.scrollbar-thumb-blue::-webkit-scrollbar-thumb {
  --bg-opacity: 1;
  background-color: #edf2f7;
  background-color: rgba(237, 242, 247, var(--bg-opacity));
}

.scrollbar-thumb-rounded::-webkit-scrollbar-thumb {
  border-radius: 0.25rem;
}
</style>
