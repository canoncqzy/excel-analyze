import * as ExcelJS from 'exceljs'
// import { FillPattern } from 'exceljs/index'

export default class PayingStudents {
  remarkColumn: readonly ExcelJS.CellValue[]
  studentColumn:readonly ExcelJS.CellValue[]
  workbook: ExcelJS.Workbook
  worksheet: ExcelJS.Worksheet
  worksheetErr: ExcelJS.Worksheet
  studentErr: Set<string>

  constructor (remarkColumn: readonly ExcelJS.CellValue[], studentColumn: readonly ExcelJS.CellValue[]) {
    this.remarkColumn = remarkColumn
    this.studentColumn = studentColumn
    this.workbook = new ExcelJS.Workbook()
    this.worksheet = this.workbook.addWorksheet('学生名单')
    this.worksheetErr = this.workbook.addWorksheet('重名学生名单')
    this.studentErr = new Set([])
  }

  /**
   * 设置背景景色配置
   * @param color
   */
  setCellFill (color: string): ExcelJS.FillPattern {
    return {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {
        argb: color
      }
    }
  }

  /**
   * 设置边框样式
   */
  setBorderColor (): Partial<ExcelJS.Borders> {
    const border: ExcelJS.Border = {
      style: 'thin',
      color: {
        argb: 'FFC3CBDD'
      }
    }
    return {
      top: border,
      left: border,
      bottom: border,
      right: border,
      diagonal: {
        up: true,
        down: true
      }
    }
  }

  /**
   * 是否是重复学生
   * @param index
   * @param student
   */
  judgeSameName (index: number, student: string): boolean {
    for (let j = 1; j < this.studentColumn.length; j++) {
      if (index !== j) {
        const other = this.studentColumn[j] as string
        const repeat:boolean = other.includes(student)
        if (repeat) {
          this.studentErr.add(student)
          return repeat
        }
      }
    }
    return false
  }

  /**
   * 添加付款学生
   * @param student
   */
  setStudentsPaid (student: string) {
    for (let i = 1; i < this.remarkColumn.length; i++) {
      const comment = this.remarkColumn[i] as string
      const commentCell = this.worksheet.getCell(i, 1)
      commentCell.value || (commentCell.value = comment)
      const repeat:boolean = comment.includes(student)
      if (repeat) {
        commentCell.fill = this.setCellFill('FFA9D08E')
        return repeat
      }
    }
    return false
  }

  async start () {
    for (let i = 1; i < this.studentColumn.length; i++) {
      const student = this.studentColumn[i] as string
      const studentCell = this.worksheet.getCell(i, 2)
      studentCell.value = student
      const repeat = this.judgeSameName(i, student)
      if (repeat) {
        studentCell.fill = this.setCellFill('FFFF0000')
      }
      const paid = repeat ? false : this.setStudentsPaid(student)
      if (paid) {
        const paidCell = this.worksheet.getCell(i, 3)
        paidCell.value = 1
      }
    }
    this.worksheet.insertRow(1, ['留言信息', '学生名单', '付费情况'])
    this.worksheet.getRow(1).eachCell((cell) => {
      cell.style = {
        alignment: {
          horizontal: 'center',
          vertical: 'middle'
        },
        border: this.setBorderColor(),
        font: {
          size: 12
        },
        fill: this.setCellFill('FFFFFF00')
      }
    })
    this.studentErr.forEach((value:string) => {
      this.worksheetErr.addRow([value])
    })
    const buffer = await this.workbook.xlsx.writeBuffer()
    this.writeFile('分析结果', buffer)
  }

  /**
   * 导出文件
   * @param fileName
   * @param content
   */
  writeFile (fileName: string, content: ExcelJS.Buffer) {
    const a = document.createElement('a')
    const blob = new Blob([content])
    a.download = `${fileName}.xlsx`
    a.href = URL.createObjectURL(blob)
    a.click()
  }
}
