package io.luxus.dexcel

import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.data.forAll
import io.kotest.data.headers
import io.kotest.data.row
import io.kotest.data.table
import io.kotest.matchers.comparables.shouldBeGreaterThanOrEqualTo
import io.kotest.matchers.comparables.shouldBeLessThanOrEqualTo
import io.kotest.matchers.equals.shouldBeEqual
import io.kotest.matchers.shouldBe
import io.kotest.matchers.shouldNotBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File

/**
 * @author kjkim
 * @since 2025. 3. 1.
 */
class ExcelExtensionsKtTest: DescribeSpec({
    describe("excelColumnIndex") {
        table(
            headers("input", "expected"),
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

    describe("excelColumnName") {
        table(
            headers("input", "expected"),
            row(0, "A"),
            row(1, "B"),
            row(25, "Z"),
            row(26, "AA"),
            row(27, "AB"),
            row(51, "AZ"),
        ).forAll { input, expected ->
            val result = input.excelColumnName
            it("excelColumnName of $input should be $expected") {
                result shouldBeEqual expected
            }
        }

        table(
            headers("invalidInput"),
            row(-1), // negative
        ).forAll { invalidInput ->
            it("$invalidInput should throw IllegalArgumentException") {
                val exception = shouldThrow<IllegalArgumentException> {
                    invalidInput.excelColumnName
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

        context("sheet is empty"){
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

        context("row is empty") {
            table(
                headers("nullCellName"),
                row("A1"),
                row("B1"),
                row("AAA1"),
                row("ZZZAAA1"),
            ).forAll { nullCellName ->
                it("$nullCellName should return null") {
                    // given
                    val workbook = XSSFWorkbook()
                    val sheet = workbook.createSheet("sheet1")
                    sheet.createRow(0)

                    // when
                    val result = sheet.getCell(nullCellName)

                    // then
                    result shouldBe null
                }
            }
        }

        context("no row name") {
            table(
                headers("nullCellName"),
                row("A"),
                row("Z"),
                row("AA"),
                row("AZ"),
                row("AZCBDE"),
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

        context("no cell name") {
            table(
                headers("nullCellName"),
                row("1"),
                row("123"),
                row("123123"),
                row("12222789"),
                row("123123123"),
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
    }

    val pixelToWidth = mutableMapOf<Int, Int>()
    val poi256WidthToWidthList = mutableListOf<Pair<Double, Int>>()
    XSSFWorkbook(File("./src/test/resources/ExcelColumnWidth_PoiFont.xlsx").inputStream()).read {
        sheet("sheet1") {
            row(0) {
                for (i in 0..75) {
                    pixelToWidth[i] = sheet.getColumnWidth(i)
                    poi256WidthToWidthList += (double(i)!! to sheet.getColumnWidth(i))
                }
            }
        }
    }

    describe("poiPixelToWidth") {
        table(
            headers("input", "minimumValue"),
            *(0..75).map { row(it, pixelToWidth[it]!!) }.toTypedArray()
        ).forAll { input, minimumValue ->
            val result = input.poiPixelToWidth

            it("poi256Width of $input should be between $minimumValue and ${minimumValue + 5}") {
                result shouldBeGreaterThanOrEqualTo minimumValue
                result shouldBeLessThanOrEqualTo minimumValue + 5
            }
        }
    }

    describe("poiWidthToPoi256Width") {
        table(
            headers("input", "minimumValue"),
            *(0..75).map { row(
                poi256WidthToWidthList[it].first,
                poi256WidthToWidthList[it].second
            ) }.toTypedArray()
        ).forAll { input, minimumValue ->
            val result = input.poiWidthToPoi256Width

            it("poiWidthToPoi256Width of $input should be between $minimumValue and ${minimumValue + 5}") {
                result shouldBeGreaterThanOrEqualTo minimumValue
                result shouldBeLessThanOrEqualTo minimumValue + 5
            }
        }
    }

    // excel(File("./src/test/resources/ExcelColumnWidth.xlsx").outputStream()) {
    //     sheet("pixelToWidth") {
    //         for (i in 0..75) {
    //             row {
    //                 cell(i.toDouble())
    //                 cell(pixelToWidth[i]!!.toDouble())
    //             }
    //         }
    //     }
    //
    //     sheet("poi256WidthToWidth") {
    //         poi256WidthToWidthList.forEach { (poi256Width, width) ->
    //             row {
    //                 cell(poi256Width)
    //                 cell(width.toDouble())
    //             }
    //         }
    //     }
    // }
})
