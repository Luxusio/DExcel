package io.luxus.dexcel

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.equals.shouldBeEqual
import io.mockk.every
import io.mockk.mockk
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Assertions.*
import java.io.ByteArrayOutputStream

/**
 * @author luxus
 * @since 2025. 2. 28.
 */
class DExcelKtTest: DescribeSpec({
    describe("excel write") {
        it("should be call functions in valid order") {
            // given
            val outputStream = ByteArrayOutputStream()
            var callOrder = 0
            val plugin = mockk<ExcelWriterPlugin>()
            every { plugin.beforeWorkbook(any()) } answers {
                callOrder shouldBeEqual 0
                callOrder = 1
            }
            every { plugin.afterWorkbook(any()) } answers {
                callOrder shouldBeEqual 2
                callOrder = 3
            }

            // when
            excel(outputStream, plugins = listOf(plugin)) {
                callOrder shouldBeEqual 1
                callOrder = 2
            }

            // then
            callOrder shouldBeEqual 3
        }

        it("outputStream is valid excel file") {
            // given
            val outputStream = ByteArrayOutputStream()

            // when
            excel(outputStream) {
                sheet("sheet1") {
                    row {
                        cell("A1")
                        cell("B1")
                    }
                    row {
                        cell("A2")
                        cell("B2")
                    }
                }
            }

            // then
            val workbook = XSSFWorkbook(outputStream.toByteArray().inputStream())
            val sheet = workbook.getSheetAt(0)
            assertEquals("A1", sheet.getRow(0).getCell(0).stringCellValue)
            assertEquals("B1", sheet.getRow(0).getCell(1).stringCellValue)
            assertEquals("A2", sheet.getRow(1).getCell(0).stringCellValue)
            assertEquals("B2", sheet.getRow(1).getCell(1).stringCellValue)
        }
    }

    describe("Workbook.read") {
        val workbook = XSSFWorkbook()
        excel(ByteArrayOutputStream(), workbook) {
            sheet("sheet1") {
                row {
                    cell("A1")
                    cell("B1")
                }
                row {
                    cell("A2")
                    cell("B2")
                }
            }
        }

        it("should read workbook") {
            // when
            val result = workbook.read {
                sheet("sheet1") {
                    rows {
                        strings()
                    }
                }
            }.toList()

            // then
            assertEquals(listOf(listOf("A1", "B1"), listOf("A2", "B2")), result)
        }
    }
})
