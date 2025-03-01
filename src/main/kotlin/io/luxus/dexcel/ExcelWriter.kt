package io.luxus.dexcel

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import java.io.OutputStream
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.Calendar
import java.util.Date

interface ExcelWriterPlugin {
    fun beforeWorkbook(writer: ExcelWriter) {}
    fun afterWorkbook(writer: ExcelWriter) {}
    fun beforeSheet(writer: SheetWriter) {}
    fun afterSheet(writer: SheetWriter) {}
    fun beforeRow(writer: RowWriter) {}
    fun afterRow(writer: RowWriter) {}
}

class ExcelWriter(
    val workbook: Workbook,
    val plugins: List<ExcelWriterPlugin>,
) : AutoCloseable {
    fun sheet(name: String, block: SheetWriter.() -> Unit) {
        val sheetWriter = SheetWriter(workbook.createSheet(name), plugins)
        plugins.forEach { it.beforeSheet(sheetWriter) }
        block(sheetWriter)
        plugins.forEach { it.afterSheet(sheetWriter) }
    }

    fun write(outputStream: OutputStream) {
        workbook.write(outputStream)
    }

    override fun close() {
        workbook.close()
    }
}

class SheetWriter(
    val sheet: Sheet,
    val plugins: List<ExcelWriterPlugin>,
) {
    var rownum = 0
        private set

    fun row(block: RowWriter.() -> Unit) {
        val rowWriter = RowWriter(sheet.createRow(rownum++))
        plugins.forEach { it.beforeRow(rowWriter) }
        block(rowWriter)
        plugins.forEach { it.afterRow(rowWriter) }
    }
}

class RowWriter(val row: Row) {
    var colnum = 0
        private set

    private fun cell(setCellValue: Cell.() -> Unit, block: Cell.() -> Unit) {
        val cell = row.createCell(colnum++)
        setCellValue(cell)
        block(cell)
    }

    fun cell(block: Cell.() -> Unit = { }) = cell({ }, block)
    fun cell(value: Double?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(it) } }, block)
    fun cell(value: Date?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(it) } }, block)
    fun cell(value: LocalDate?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(it) } }, block)
    fun cell(value: LocalDateTime?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(it) } }, block)
    fun cell(value: Calendar?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(value) } }, block)
    fun cell(value: RichTextString?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(it) } }, block)
    fun cell(value: String?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(it) } }, block)
    fun cell(value: Boolean?, block: Cell.() -> Unit = { }) = cell({ value?.let { setCellValue(it) } }, block)
}
