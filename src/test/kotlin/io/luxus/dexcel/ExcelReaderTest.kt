package io.luxus.dexcel

import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.data.forAll
import io.kotest.data.headers
import io.kotest.data.row
import io.kotest.data.table
import io.kotest.matchers.equals.shouldBeEqual
import io.kotest.matchers.shouldBe
import io.kotest.matchers.shouldNotBe
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.lang.IllegalArgumentException

/**
 * @author kjkim
 * @since 2025. 3. 4.
 */
class ExcelReaderTest: DescribeSpec({
    describe("ExcelReader") {
        describe("sheet") {
            context("if sheet name is null") {
                val excelReader = ExcelReader(XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        row {
                            cell("A1")
                            cell("B1")
                        }
                        row {
                            cell("A2")
                            cell("B2")
                        }
                    }
                })

                var captured: SheetReader? = null
                val returned = excelReader.sheet {
                    captured = this
                    "returned result"
                }

                it("should read the first sheet") {
                    captured shouldNotBe null
                    captured!!.sheet.sheetName shouldBeEqual "sheet1"
                    captured.formatter shouldNotBe null
                }

                it("should return the result of block") {
                    returned shouldBeEqual "returned result"
                }
            }

            context("if sheet name is not null") {
                val excelReader = ExcelReader(XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        row {
                            cell("A1")
                            cell("B1")
                        }
                        row {
                            cell("A2")
                            cell("B2")
                        }
                    }
                })

                var captured: SheetReader? = null
                val returned = excelReader.sheet("sheet1") {
                    captured = this
                    "returned result"
                }

                it("should read the sheet") {
                    captured shouldNotBe null
                    captured!!.sheet.sheetName shouldBeEqual "sheet1"
                    captured.formatter shouldNotBe null
                }

                it("should return the result of block") {
                    returned shouldBeEqual "returned result"
                }
            }

            context("if sheet name is invalid") {
                val excelReader = ExcelReader(XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        row {
                            cell("A1")
                            cell("B1")
                        }
                        row {
                            cell("A2")
                            cell("B2")
                        }
                    }
                })

                it("should throw IllegalArgumentException") {
                    shouldThrow<IllegalArgumentException> {
                        excelReader.sheet("invalid") { }
                    }
                }
            }

            context("if sheet name is null and workbook is empty") {
                val excelReader = ExcelReader(XSSFWorkbook())

                it("should throw IllegalArgumentException") {
                    val result = shouldThrow<IllegalArgumentException> {
                        excelReader.sheet { }
                    }
                    result.message shouldBe "Sheet not found"
                }
            }
        }

        describe("close") {
            it("should close workbook") {
                val workbook = XSSFWorkbook()
                val excelReader = ExcelReader(workbook)
                excelReader.close()
                shouldThrow<IllegalArgumentException> {
                    workbook.getSheetAt(0)
                }
            }
        }
    }

    describe("SheetReader") {
        describe("rows") {
            context("with no arguments") {
                val workbook = XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        for (i in 1..20) {
                            row {
                                cell("A$i")
                            }
                        }
                    }
                }.getSheet("sheet1")

                val sheetReader = SheetReader(workbook, DataFormatter())
                val list = sheetReader.rows { this }.toList()

                it("should return sequence that contains all rows") {
                    list.size shouldBe 20
                }

                it("each element should called with different row") {
                    list.forEachIndexed { index, rowReader ->
                        rowReader.row shouldBe workbook.getRow(index)
                    }
                }
            }

            context("with startRow") {
                val workbook = XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        for (i in 1..20) {
                            row {
                                cell("A$i")
                            }
                        }
                    }
                }.getSheet("sheet1")

                val sheetReader = SheetReader(workbook, DataFormatter())
                val sequence = sheetReader.rows(5) { }

                it("should return sequence that contains rows from startRow") {
                    sequence.toList().size shouldBe 15
                }
            }

            context("with endRow") {
                val workbook = XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        for (i in 1..20) {
                            row {
                                cell("A$i")
                            }
                        }
                    }
                }.getSheet("sheet1")

                val sheetReader = SheetReader(workbook, DataFormatter())
                val sequence = sheetReader.rows(endRow = 4) { }

                it("should return sequence that contains rows until endRow") {
                    sequence.toList().size shouldBe 5
                }
            }

            context("with startRow and endRow") {
                val workbook = XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        for (i in 1..20) {
                            row {
                                cell("A$i")
                            }
                        }
                    }
                }.getSheet("sheet1")

                val sheetReader = SheetReader(workbook, DataFormatter())
                val sequence = sheetReader.rows(5, 10) { }

                it("should return sequence that contains rows from startRow to endRow") {
                    sequence.toList().size shouldBe 6
                }
            }

            context("with invalid startRow") {
                val workbook = XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        for (i in 1..20) {
                            row {
                                cell("A$i")
                            }
                        }
                    }
                }.getSheet("sheet1")

                val sheetReader = SheetReader(workbook, DataFormatter())
                val sequence = sheetReader.rows(20, 10) { }

                it("should return empty sequence") {
                    sequence.toList().size shouldBe 0
                }
            }

            context("with empty sheet") {
                val workbook = XSSFWorkbook().createSheet("sheet1")
                val sheetReader = SheetReader(workbook, DataFormatter())
                val sequence = sheetReader.rows { }

                it("should return empty sequence") {
                    sequence.toList().size shouldBe 0
                }
            }

            context("with 1 row sheet") {
                val workbook = XSSFWorkbook().apply {
                    ExcelWriter(this, listOf()).sheet("sheet1") {
                        row {
                            cell("A1")
                        }
                    }
                }.getSheet("sheet1")

                val sheetReader = SheetReader(workbook, DataFormatter())
                val sequence = sheetReader.rows { }

                it("should return sequence that contains 1 row") {
                    sequence.toList().size shouldBe 1
                }
            }
        }
    }

    describe("RowReader") {
        fun createRow(block: RowWriter.() -> Unit): Row {
            return XSSFWorkbook().apply {
                ExcelWriter(this, listOf()).sheet("sheet1") {
                    row(block)
                }
            }.getSheet("sheet1").getRow(0)
        }

        describe("strings") {
            table(
                headers("row", "value"),
                row(createRow { }, listOf()),
                row(createRow {
                    cell()
                    cell()
                }, listOf("", "")),
                row(createRow {
                    cell(777.0)
                    cell(777.0)
                    cell(777.0)
                }, listOf("777", "777", "777")),
                row(createRow {
                    cell("A1")
                    cell("B1")
                    cell(123.0)
                    cell(123.4)
                    cell(true)
                }, listOf("A1", "B1", "123", "123.4", "TRUE")),
            ).forAll { row, value ->
                context("with valid cells ($value)") {
                    val rowReader = RowReader(row, DataFormatter())
                    val list = rowReader.strings()

                    it("should return list of cell values ($value)") {
                        list shouldBe value
                    }
                }
            }
        }

        describe("string") {
            table(
                headers("row", "value"),
                row(createRow { }, ""),
                row(createRow { cell() }, ""),
                row(createRow { cell("A1") }, "A1"),
                row(createRow { cell("Hello world it's me hello.") }, "Hello world it's me hello."),
                row(createRow { cell(123.0) }, "123"),
                row(createRow { cell(123.4) }, "123.4"),
                row(createRow { cell(true) }, "TRUE"),
                row(createRow { cell(false) }, "FALSE"),
            ).forAll { row, value ->
                context("with valid cell (${row.getCell(0)})") {
                    val rowReader = RowReader(row, DataFormatter())
                    val result = rowReader.string(0)

                    it("should return cell value ($value)") {
                        result shouldBe value
                    }
                }
            }
        }

        describe("double") {
            table(
                headers("row", "value"),
                row(createRow { }, null),
                row(createRow { cell() }, null),
                row(createRow { cell(123.0) }, 123.0),
                row(createRow { cell(123.4) }, 123.4),
                row(createRow { cell("01100") }, 1100),
                row(createRow { cell("0x30") }, null),
                row(createRow { cell("0b1100") }, null),
                row(createRow { cell(true) }, null),
                row(createRow { cell(false) }, null),
            ).forAll { row, value ->
                context("with valid cell ${row.getCell(0)}") {
                    val rowReader = RowReader(row, DataFormatter())
                    val result = rowReader.double(0)

                    it("should return cell value $value") {
                        result shouldBe value
                    }
                }
            }
        }

        describe("int") {
            table(
                headers("row", "value"),
                row(createRow { }, null),
                row(createRow { cell() }, null),
                row(createRow { cell(123.0) }, 123),
                row(createRow { cell(123.4) }, null),
                row(createRow { cell("01100") }, 1100),
                row(createRow { cell("0x30") }, null),
                row(createRow { cell("0b1100") }, null),
                row(createRow { cell(true) }, null),
                row(createRow { cell(false) }, null),
            ).forAll { row, value ->
                context("with valid cell ${row.getCell(0)}") {
                    val rowReader = RowReader(row, DataFormatter())
                    val result = rowReader.int(0)

                    it("should return cell value $value") {
                        result shouldBe value
                    }
                }
            }
        }

        describe("long") {
            table(
                headers("row", "value"),
                row(createRow { }, null),
                row(createRow { cell() }, null),
                row(createRow { cell(123.0) }, 123L),
                row(createRow { cell(123.4) }, null),
                row(createRow { cell("01100") }, 1100L),
                row(createRow { cell("0x30") }, null),
                row(createRow { cell("0b1100") }, null),
                row(createRow { cell(true) }, null),
                row(createRow { cell(false) }, null),
            ).forAll { row, value ->
                context("with valid cell ${row.getCell(0)}") {
                    val rowReader = RowReader(row, DataFormatter())
                    val result = rowReader.long(0)

                    it("should return cell value $value") {
                        result shouldBe value
                    }
                }
            }
        }

        describe("boolean") {
            table(
                headers("row", "value"),
                row(createRow { }, null),
                row(createRow { cell() }, null),
                row(createRow { cell(123.0) }, null),
                row(createRow { cell(123.4) }, null),
                row(createRow { cell("01100") }, null),
                row(createRow { cell("0x30") }, null),
                row(createRow { cell("0b1100") }, null),
                row(createRow { cell(0.0) }, null),
                row(createRow { cell(1.0) }, null),
                row(createRow { cell(true) }, true),
                row(createRow { cell(false) }, false),
                row(createRow { cell("TRUE") }, true),
                row(createRow { cell("FALSE") }, false),
                row(createRow { cell("true") }, true),
                row(createRow { cell("false") }, false),
            ).forAll { row, value ->
                context("with valid cell ${row.getCell(0)}") {
                    val rowReader = RowReader(row, DataFormatter())
                    val result = rowReader.boolean(0)

                    it("should return cell value $value") {
                        result shouldBe value
                    }
                }
            }
        }
    }
})
