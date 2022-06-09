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
import type { UploadChangeParam } from 'ant-design-vue'
import * as ExcelJS from 'exceljs'

interface UploadParam extends UploadChangeParam {
  onSuccess(): void
}

const writeFile = (fileName: string, content: ExcelJS.Buffer) => {
  const a = document.createElement('a')
  const blob = new Blob([content])
  a.download = `${fileName}.xlsx`
  a.href = URL.createObjectURL(blob)
  a.click()
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

// if (i === 6) {
//   cell.fill = {
//     type: 'pattern',
//     pattern: 'solid',
//     fgColor: {
//       argb: 'FFA9D08E'
//     }
//   }
// }
const handleData = async (remarkColumn: readonly ExcelJS.CellValue[], studentColumn: readonly ExcelJS.CellValue[]) => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('学生名单')
  /**
   * 查重
   */
  for (let i = 1; i < studentColumn.length; i++) {
    const student = studentColumn[i]
    const cell = worksheet.getCell(i, 1)
    cell.value = student
    for (let j = 1; j < studentColumn.length; j++) {
      if (i !== j) {
        const other = studentColumn[j]
        const repeat = (other as string).includes(student as string)
        if (repeat) {
          const cell = worksheet.getCell(i, 1)
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb: '7FF56C6C'
            }
          }
        }
      }
    }
  }
  const firstColumn = worksheet.getColumn(1)
  firstColumn.header = '学生'
  const buffer = await workbook.xlsx.writeBuffer()
  writeFile('分析结果', buffer)
}

const customRequest = async (uploadChangeParam: UploadParam) => {
  const file: any = uploadChangeParam.file
  const workbook = await parseFile(file)
  const workSheet = workbook.getWorksheet(1)
  const remarkColumn = workSheet.getColumn(1).values
  const studentColumn = workSheet.getColumn(2).values
  await handleData(remarkColumn, studentColumn)
  uploadChangeParam.onSuccess()
}

// const addWorkbook = (rowList: string[]) => {
//   // 创建工作簿
//   const workbook = new ExcelJS.Workbook()
//   // 添加工作表
//   const worksheet = workbook.addWorksheet('sheet1')
// }

</script>

<style scoped lang="less">
.container {
  margin: 150px auto;
  width: 45%;
}
</style>
