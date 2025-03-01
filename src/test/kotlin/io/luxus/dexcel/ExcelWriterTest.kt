package io.luxus.dexcel

import io.kotest.assertions.throwables.shouldNotThrowAny
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.data.forAll
import io.kotest.data.headers
import io.kotest.data.row
import io.kotest.data.table
import io.kotest.matchers.equals.shouldBeEqual
import io.kotest.matchers.shouldBe
import io.kotest.matchers.shouldNotBe
import io.mockk.every
import io.mockk.mockk
import io.mockk.verify
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayOutputStream
import java.io.File
import java.text.SimpleDateFormat
import java.time.LocalDate
import java.util.*


/**
 * @author kjkim
 * @since 2025. 3. 1.
 */
class ExcelWriterTest: DescribeSpec({
    describe("ExcelWriter") {
        describe("sheet") {
            it("should be call functions in valid order") {
                // given
                val workbook = XSSFWorkbook()
                var callOrder = 0
                val plugin = mockk<ExcelWriterPlugin>()
                every { plugin.beforeSheet(any()) } answers {
                    callOrder shouldBeEqual 0
                    callOrder = 1
                }
                every { plugin.afterSheet(any()) } answers {
                    callOrder shouldBeEqual 2
                    callOrder = 3
                }

                // when
                ExcelWriter(workbook, listOf(plugin)).sheet("sheetName") {
                    callOrder shouldBeEqual 1
                    callOrder = 2
                }

                // then
                callOrder shouldBeEqual 3
            }

            it("should create sheet") {
                // given
                val workbook = XSSFWorkbook()
                val sheetName = "sheetName"

                // when
                ExcelWriter(workbook, listOf()).sheet(sheetName) {
                    // do nothing
                }

                // then
                workbook.getSheet(sheetName) shouldNotBe null
            }
        }

        describe("write") {
            it("should write workbook to output stream") {
                // given
                val workbook = XSSFWorkbook().apply {
                    createSheet("sheet1").apply {
                        createRow(0).apply {
                            createCell(0).setCellValue("test")
                        }
                    }
                }

                val outputStream = ByteArrayOutputStream()

                // when
                ExcelWriter(workbook, listOf()).write(outputStream)

                // then
                outputStream.toByteArray().size shouldNotBe 0
                val resultWorkbook = XSSFWorkbook(outputStream.toByteArray().inputStream())
                resultWorkbook.getSheetAt(0).getRow(0).getCell(0).stringCellValue shouldBeEqual "test"
            }
        }

        describe("close") {
            it("should close workbook") {
                // given
                val workbook = mockk<Workbook>()
                every { workbook.close() } returns Unit
                val excelWriter = ExcelWriter(workbook, listOf())

                // when
                excelWriter.close()

                // then
                verify(exactly = 1) { workbook.close() }
            }
        }
    }

    describe("SheetWriter") {
        describe("row") {
            it("should be call functions in valid order") {
                // given
                val sheet = XSSFWorkbook().createSheet("sheet1")
                var callOrder = 0
                val plugin = mockk<ExcelWriterPlugin>()
                every { plugin.beforeRow(any()) } answers {
                    callOrder shouldBeEqual 0
                    callOrder = 1
                }
                every { plugin.afterRow(any()) } answers {
                    callOrder shouldBeEqual 2
                    callOrder = 3
                }

                // when
                SheetWriter(sheet, listOf(plugin)).row {
                    callOrder shouldBeEqual 1
                    callOrder = 2
                }

                // then
                callOrder shouldBeEqual 3
            }

            table(
                headers("n"),
                row(0),
                row(1),
                row(2),
                row(3),
                row(9),
            ).forAll { n ->
                it("should create row $n times") {
                    // given
                    val sheet = XSSFWorkbook().createSheet("sheet1")
                    val sheetWriter = SheetWriter(sheet, listOf())

                    // when
                    for (i in 0 until n) {
                        sheetWriter.row {
                            // do nothing
                        }
                    }

                    // then
                    sheetWriter.rownum shouldBe n
                    for (i in 0 until n) {
                        sheet.getRow(i) shouldNotBe null
                    }
                    for (i in n until 10) {
                        sheet.getRow(i) shouldBe null
                    }
                }
            }
        }
    }

    describe("RowWriter") {
        describe("cell") {
            it("should be call functions in valid order") {
                // given
                val row = XSSFWorkbook().createSheet("sheet1").createRow(0)
                var callOrder = 0
                val rowWriter = RowWriter(row)

                // when
                rowWriter.cell {
                    callOrder shouldBeEqual 0
                    callOrder = 1
                }

                // then
                callOrder shouldBeEqual 1
            }

            table(
                headers("n"),
                row(0),
                row(1),
                row(2),
                row(3),
                row(9),
            ).forAll { n ->
                it("should create cell $n times") {
                    // given
                    val row = XSSFWorkbook().createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    // when
                    for (i in 0 until n) {
                        rowWriter.cell {
                            // do nothing
                        }
                    }

                    // then
                    rowWriter.colnum shouldBe n
                    for (i in 0 until n) {
                        row.getCell(i) shouldNotBe null
                    }
                    for (i in n until 10) {
                        row.getCell(i) shouldBe null
                    }
                }
            }
        }

        describe("cell(Double?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row(1.0, "1"),
                row(1.1, "1.1"),
                row(1.123456789, "1.123456789"),
            ).forAll { value, expected ->
                it("should set cell value with double") {
                    // given
                    val row = XSSFWorkbook().createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)
                    val formatter = DataFormatter()

                    // when
                    rowWriter.cell(value)

                    // then
                    row.getCell(0).numericCellValue shouldBeEqual (value ?: 0.0)
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }

        describe("cell(Date?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row(Date(0), "1970-01-01 00:00:00"),
                row(Date(1), "1970-01-01 00:00:00"),
                row(Date(1000), "1970-01-01 00:00:01"),
            ).forAll { value, expected ->
                it("should set cell value with date") {
                    // given
                    val workbook = XSSFWorkbook()
                    val row = workbook.createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    val creationHelper = workbook.creationHelper
                    val dateStyle = workbook.createCellStyle().apply {
                        dataFormat = creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss")
                    }

                    // when
                    rowWriter.cell(value) {
                        cellStyle = dateStyle
                    }

                    // then
                    val formatter = DataFormatter().apply {
                        addFormat("yyyy-MM-dd HH:mm:ss", SimpleDateFormat("yyyy-MM-dd HH:mm:ss").apply {
                            timeZone = TimeZone.getTimeZone("UTC")
                        })
                    }
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }

        describe("cell(LocalDate?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row(LocalDate.of(1971, 11, 21), "1971-11-21"),
                row(LocalDate.of(2025, 3, 1), "2025-03-01"),
            ).forAll { value, expected ->
                it("should set cell value with local date") {
                    // given
                    val workbook = XSSFWorkbook()
                    val row = workbook.createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    val creationHelper = workbook.creationHelper
                    val dateStyle = workbook.createCellStyle().apply {
                        dataFormat = creationHelper.createDataFormat().getFormat("yyyy-MM-dd")
                    }

                    // when
                    rowWriter.cell(value) {
                        cellStyle = dateStyle
                    }

                    // then
                    val formatter = DataFormatter().apply {
                        addFormat("yyyy-MM-dd", SimpleDateFormat("yyyy-MM-dd"))
                    }
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }

        describe("cell(LocalDateTime?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row(LocalDate.of(1971, 11, 21).atStartOfDay(), "1971-11-21 00:00:00"),
                row(LocalDate.of(2025, 3, 1).atStartOfDay(), "2025-03-01 00:00:00"),
            ).forAll { value, expected ->
                it("should set cell value with local date time") {
                    // given
                    val workbook = XSSFWorkbook()
                    val row = workbook.createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    val creationHelper = workbook.creationHelper
                    val dateStyle = workbook.createCellStyle().apply {
                        dataFormat = creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss")
                    }

                    // when
                    rowWriter.cell(value) {
                        cellStyle = dateStyle
                    }

                    // then
                    val formatter = DataFormatter().apply {
                        addFormat("yyyy-MM-dd HH:mm:ss", SimpleDateFormat("yyyy-MM-dd HH:mm:ss"))
                    }
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }

        describe("cell(Calendar?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row(Calendar.getInstance().apply { timeInMillis = 0 }, "1970-01-01 00:00:00"),
                row(Calendar.getInstance().apply { timeInMillis = 1 }, "1970-01-01 00:00:00"),
                row(Calendar.getInstance().apply { timeInMillis = 1000 }, "1970-01-01 00:00:01"),
            ).forAll { value, expected ->
                it("should set cell value with calendar") {
                    // given
                    val workbook = XSSFWorkbook()
                    val row = workbook.createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    val creationHelper = workbook.creationHelper
                    val dateStyle = workbook.createCellStyle().apply {
                        dataFormat = creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss")
                    }

                    // when
                    rowWriter.cell(value) {
                        cellStyle = dateStyle
                    }

                    // then
                    val formatter = DataFormatter().apply {
                        addFormat("yyyy-MM-dd HH:mm:ss", SimpleDateFormat("yyyy-MM-dd HH:mm:ss").apply {
                            timeZone = TimeZone.getTimeZone("UTC")
                        })
                    }
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }

        describe("cell(RichTextString?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row(XSSFRichTextString(null as String?), ""),
                row(XSSFRichTextString("test"), "test"),
            ).forAll { value, expected ->
                it("should set cell value with rich text string") {
                    // given
                    val workbook = XSSFWorkbook()
                    val row = workbook.createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    // when
                    rowWriter.cell(value)

                    // then
                    val formatter = DataFormatter()
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }

        describe("cell(String?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row("", ""),
                row("test", "test"),
            ).forAll { value, expected ->
                it("should set cell value with string") {
                    // given
                    val workbook = XSSFWorkbook()
                    val row = workbook.createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    // when
                    rowWriter.cell(value)

                    // then
                    val formatter = DataFormatter()
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }

        describe("cell(Boolean?)") {
            table(
                headers("value", "expected"),
                row(null, ""),
                row(true, "TRUE"),
                row(false, "FALSE"),
            ).forAll { value, expected ->
                it("should set cell value with boolean") {
                    // given
                    val workbook = XSSFWorkbook()
                    val row = workbook.createSheet("sheet1").createRow(0)
                    val rowWriter = RowWriter(row)

                    // when
                    rowWriter.cell(value)

                    // then
                    val formatter = DataFormatter()
                    formatter.formatCellValue(row.getCell(0)) shouldBeEqual expected
                }
            }
        }
    }

    describe("ExcelWriterPlugin") {
        describe("default functions called without errors") {
            // given
            val plugin = object: ExcelWriterPlugin {}

            // when, then
            shouldNotThrowAny {
                ExcelWriter(XSSFWorkbook(), listOf(plugin)).sheet("sheet1") {
                    row {
                        cell("A1")
                        cell("B1")
                    }
                }
            }
        }
    }
})
