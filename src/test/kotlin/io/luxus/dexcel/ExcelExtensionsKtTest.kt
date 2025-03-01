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
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * @author kjkim
 * @since 2025. 3. 1.
 */
class ExcelExtensionsKtTest: DescribeSpec({
    describe("excelColumnIndex") {
        table(headers("input", "expected"),
            row("A", 0),
            row("B", 1),
            row("Z", 25),
            row("AA", 26),
            row("AB", 27),
            row("AZ", 51),
        ).forAll { input, expected ->
            val result = input.excelColumnIndex
            it("excelColumnIndex of $input should be $expected") {
                result shouldBeEqual expected
            }
        }

        table(
            headers("invalidInput"),
            row(""), // empty
            row("a"), // lower case
            row("1"), // number
            row("AA1"), // number
            row("A1"), // number
            row("A1A"), // number
        ).forAll { invalidInput ->
            it("$invalidInput should throw IllegalArgumentException") {
                val exception = shouldThrow<IllegalArgumentException> {
                    invalidInput.excelColumnIndex
                }
                exception.message shouldNotBe null
                exception.message!! shouldBeEqual "invalid input: $invalidInput"
            }
        }
    }

    describe("Sheet.getCell") {
        it("should return cell") {
            // given
            val workbook = XSSFWorkbook()
            val sheet = workbook.createSheet("sheet1")
            val row = sheet.createRow(0)
            val cell = row.createCell(0)
            cell.setCellValue("test")

            // when
            val result = sheet.getCell("A1")

            // then
            result shouldNotBe null
            result!!.stringCellValue shouldBeEqual "test"
        }

        table(
            headers("nullCellName"),
            row(""),
            row("A1"),
            row("A2"),
            row("B123123123"),
            row("A123123123"),
            row("ZZZZZZZ12333445"),
        ).forAll { nullCellName ->
            it("$nullCellName should return null") {
                // given
                val workbook = XSSFWorkbook()
                val sheet = workbook.createSheet("sheet1")

                // when
                val result = sheet.getCell(nullCellName)

                // then
                result shouldBe null
            }
        }
    }
})
