import * as ExcelJS from 'exceljs'
// import { FillPattern } from 'exceljs/index'

export default class PayingStudents {
  remarkColumn: readonly ExcelJS.CellValue[]
  studentColumn:readonly ExcelJS.CellValue[]
  workbook: ExcelJS.Workbook
  worksheet: ExcelJS.Worksheet

  constructor (remarkColumn: readonly ExcelJS.CellValue[], studentColumn: readonly ExcelJS.CellValue[]) {
    this.remarkColumn = remarkColumn
    this.studentColumn = studentColumn
    this.workbook = new ExcelJS.Workbook()
    this.worksheet = this.workbook.addWorksheet('学生名单')
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
        if (repeat) return repeat
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
        commentCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'FFA9D08E'
          }
        }
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
        studentCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'FFFF0000'
          }
        }
      }
      const paid = repeat ? false : this.setStudentsPaid(student)
      if (paid) {
        const paidCell = this.worksheet.getCell(i, 3)
        paidCell.value = 1
      }
    }
    this.worksheet.insertRow(1, ['留言', '学生名单', '付费情况'])
    this.worksheet.getColumn(1).style = {
      alignment: {
        horizontal: 'center',
        vertical: 'middle'
      },
      font: {
        size: 14
      },
      border: {

      },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FFFF0000'
        }
      }
    }

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
