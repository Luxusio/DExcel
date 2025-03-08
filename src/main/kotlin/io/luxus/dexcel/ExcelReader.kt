package io.luxus.dexcel

import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import java.util.Locale

class ExcelReader(
    val workbook: Workbook,
): AutoCloseable {

    private val formatter = DataFormatter()

    /**
     * Read sheet
     *
     * @param name sheet name to read. If null, read the first sheet.
     */
    fun <T> sheet(name: String? = null, block: SheetReader.() -> T): T {
        return SheetReader(getSheet(name), formatter).block()
    }

    private fun getSheet(sheetName: String?) = if (workbook.numberOfSheets == 0) {
        throw IllegalArgumentException("Sheet not found")
    } else if (sheetName == null) {
        workbook.getSheetAt(0)
    } else {
        workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Sheet '$sheetName' not found")
    }

    override fun close() {
        workbook.close()
    }
}

class SheetReader(
    val sheet: Sheet,
    val formatter: DataFormatter,
) {
    fun <T> rows(startRow: Int = sheet.firstRowNum, endRow: Int = sheet.lastRowNum, block: RowReader.() -> T): Sequence<T> {
        return sequence {
            if (endRow < 0) return@sequence
            for (rownum in startRow..endRow) {
                yield(row(rownum, block))
            }
        }
    }

    fun <T> row(rowNum: Int, block: RowReader.() -> T): T {
        val row = sheet.getRow(rowNum) ?: throw IllegalArgumentException("Row $rowNum not found")
        return RowReader(row, formatter).block()
    }
}

class RowReader(
    val row: Row,
    val formatter: DataFormatter,
) {
    fun boolean(cellNum: Int): Boolean? = string(cellNum).lowercase().toBooleanStrictOrNull()
    fun int(cellNum: Int): Int? = string(cellNum).toIntOrNull()
    fun long(cellNum: Int): Long? = string(cellNum).toLongOrNull()
    fun double(cellNum: Int): Double? = string(cellNum).toDoubleOrNull()
    fun string(cellNum: Int): String = formatter.formatCellValue(row.getCell(cellNum))
    fun strings(): List<String?> = (row.firstCellNum until row.lastCellNum).map { string(it) }
}
