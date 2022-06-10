<template>
  <div class="container">
    <a-upload-dragger
      accept=".xlsx,.xls"
      :customRequest="customRequest"
    >
      <p class="ant-upload-drag-icon">
        <inbox-outlined></inbox-outlined>
      </p>
      <p class="ant-upload-text">Click or drag file to this area to upload</p>
      <p class="ant-upload-hint">
        Support for a single or bulk upload. Strictly prohibit from uploading company data or other
        band files
      </p>
    </a-upload-dragger>
  </div>
</template>

<script setup lang="ts">
import { InboxOutlined } from '@ant-design/icons-vue'
// import type { UploadChangeParam } from 'ant-design-vue'
import * as ExcelJS from 'exceljs'
import PayingStudents from './util'

interface UploadParam {
  onSuccess(): void
  file: File
}

/**
 * 解析文件
 * @param file
 */
const parseFile: (file: File) => Promise<ExcelJS.Workbook> = (file) => {
  const workbook = new ExcelJS.Workbook()
  return new Promise(resolve => {
    const render = new FileReader()
    render.readAsArrayBuffer(file)
    render.onload = () => {
      const data = render.result as ArrayBuffer
      workbook.xlsx.load(data).then(workbook => {
        resolve(workbook)
      })
    }
  })
}

const customRequest = async (uploadChangeParam: UploadParam) => {
  const file = uploadChangeParam.file
  const workbook = await parseFile(file)
  const workSheet = workbook.getWorksheet(1)
  const remarkColumn = workSheet.getColumn(1).values
  const studentColumn = workSheet.getColumn(2).values
  const payingStudents = new PayingStudents(remarkColumn, studentColumn)
  await payingStudents.start()
  uploadChangeParam.onSuccess()
}

</script>

<style scoped lang="less">
.container {
  margin: 150px auto;
  width: 45%;
}
</style>
