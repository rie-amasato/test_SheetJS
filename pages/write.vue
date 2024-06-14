<template>
<div>
    index
    {{data}}
</div>
</template>

<script setup>
const data=ref("")

import * as ExcelJS from 'exceljs'

onMounted(()=>{
    open_excel()
})

const open_excel=async()=>{
    const path_exceltemplate="/template.xlsx"
    const file=await(await fetch(path_exceltemplate)).arrayBuffer()
    console.log(file)
    
    const data_u8array=new Uint8Array(file)
    console.log(data_u8array)

    const workbook=new ExcelJS.Workbook()
    console.log(workbook)
    
    await workbook.xlsx.load(data_u8array)
    console.log(workbook)

    const sheet=workbook.worksheets[0]
    console.log(workbook.worksheets[0].name)
    
    sheet.getCell("C3").value="かきこんだ！"
    //sheet.getCell("C3").border={
    //    top: {style: "thin", color: {argb: "000000"}},
    //    right: {style: "thin", color: {argb: "000000"}},
    //    bottom: {style: "thin", color: {argb: "000000"}},
    //    left: {style: "thin", color: {argb: "000000"}},
    //}
    
    const buffer=await workbook.xlsx.writeBuffer()
    console.log("buffer", buffer)

    const blob=new Blob([buffer], {type: "application/octet-binary"})
    console.log("blob")

    const url_download=window.URL.createObjectURL(blob)
    console.log("url", url_download)

    const a=document.createElement("a")
    a.href=url_download
    a.download="ファイルめい.xlsx"
    a.click()
    a.remove()
}

</script>