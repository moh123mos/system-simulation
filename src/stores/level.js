import { ref } from 'vue'
import { defineStore } from 'pinia'

export const useLevelStore = defineStore('level', () => {
  const level = ref('');
  const setLevel = function(lv){
    console.log(lv);
    level.value= lv;
  }
  return { level,setLevel }
})
