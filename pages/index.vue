<template>
<div>
    index
    {{data}}
</div>
</template>

<script setup lang="ts">
const data=ref("")

import * as XLSX from 'xlsx'

onMounted(()=>{
    open_excel()
})

const open_excel=async()=>{
    const path_exceltemplate="/template.xlsx"
    const file=await(await fetch(path_exceltemplate)).arrayBuffer()

    const workbook=XLSX.read(file)

    const sheet_first=workbook.Sheets[workbook.SheetNames[0]]
    data.value=(sheet_first)
}

</script>