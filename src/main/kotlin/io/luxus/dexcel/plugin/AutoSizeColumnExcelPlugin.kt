package io.luxus.dexcel.plugin

import io.luxus.dexcel.ExcelWriterPlugin
import io.luxus.dexcel.RowWriter
import io.luxus.dexcel.SheetWriter
import io.luxus.dexcel.poiPixelToWidth
import org.apache.poi.xssf.streaming.SXSSFSheet

/**
 * @author kjkim
 * @since 2025. 3. 8.
 */
class AutoSizeColumnExcelPlugin(
    val getColumns: (writer: RowWriter) -> Int? = { it.row.lastCellNum.toInt() }
) : ExcelWriterPlugin {
    private var columns: Int? = null

    override fun beforeSheet(writer: SheetWriter) {
        if (writer.sheet is SXSSFSheet) {
            writer.sheet.trackAllColumnsForAutoSizing()
        }
    }

    override fun afterRow(writer: RowWriter) {
        if (columns != null) {
            return
        }
        columns = getColumns(writer)
    }

    override fun afterSheet(writer: SheetWriter) {
        for (i in 0 until (columns ?: 0)) {
            writer.sheet.autoSizeColumn(i)
            writer.sheet.setColumnWidth(i, writer.sheet.getColumnWidth(i) + 8.poiPixelToWidth)
        }
    }
}
