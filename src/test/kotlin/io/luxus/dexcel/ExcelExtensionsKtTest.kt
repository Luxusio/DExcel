package io.luxus.dexcel

import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.data.forAll
import io.kotest.data.headers
import io.kotest.data.row
import io.kotest.data.table
import io.kotest.matchers.equals.shouldBeEqual
import io.kotest.matchers.shouldNotBe

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
            it("should throw IllegalArgumentException") {
                val exception = shouldThrow<IllegalArgumentException> {
                    invalidInput.excelColumnIndex
                }
                exception.message shouldNotBe null
                exception.message!! shouldBeEqual "invalid input: $invalidInput"
            }
        }
    }
})
