package io.luxus.dexcel

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import java.io.OutputStream

/**
 * write an Excel file to output stream
 */
fun excel(
    outputStream: OutputStream,
    workbook: Workbook = SXSSFWorkbook(),
    plugins: List<ExcelWriterPlugin> = listOf(),
    block: ExcelWriter.() -> Unit,
) {
    ExcelWriter(workbook, plugins).apply {
        plugins.forEach { it.beforeWorkbook(this) }
        block(this)
        plugins.forEach { it.afterWorkbook(this) }
        write(outputStream)
    }
}
