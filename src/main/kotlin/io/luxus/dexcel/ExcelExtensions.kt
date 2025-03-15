package io.luxus.dexcel

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import kotlin.math.pow

/**
 * changes excel word to 0-base index
 */
val String.excelColumnIndex: Int get() {
    if (this.isEmpty()) {
        throw IllegalArgumentException("invalid input: $this")
    }

    var result = 0
    val length = this.length
    for (i in 0 until length) {
        val c = this[length - 1 - i] // 역순으로 문자를 가져옵니다.
        val value = c.code - 'A'.code + 1 // 문자를 숫자로 변환합니다.
        if (value < 1 || value > 26) {
            throw IllegalArgumentException("invalid input: $this")
        }

        result += value * 26.0.pow(i.toDouble()).toInt() // 26의 거듭제곱을 적용하여 더합니다.
    }
    return result - 1 // 0-base 인덱스로 변환합니다.
}

val Int.excelColumnName: String get() {
    if (this < 0) {
        throw IllegalArgumentException("invalid input: $this")
    }

    val sb = StringBuilder()
    var n = this + 1
    while (n > 0) {
        val r = n % 26
        if (r == 0) {
            sb.append('Z')
            n = n / 26 - 1
        } else {
            sb.append('A' + r - 1)
            n /= 26
        }
    }

    return sb.reverse().toString()
}

/**
 * get cell by cell name
 * @param cellName cell name i.e. A1, B2
 */
fun Sheet.getCell(cellName: String): Cell? {
    val sb = StringBuilder()
    var i = 0
    while (i < cellName.length && cellName[i].isLetter()) {
        sb.append(cellName[i])
        i++
    }

    val column = sb.toString()
    val row = cellName.substring(i)

    if (column.isEmpty() || row.isEmpty()) {
        return null
    }

    return this.getRow(row.toInt() - 1)?.getCell(column.excelColumnIndex)
}

/**
 * @param excelWidth normal excel width
 * @return poiWidth
 */
val Int.poiPixelToWidth: Int
    get() {
        return (42.67 * this).toInt()
    }

val Double.poiWidthToPoi256Width: Int
    get() = if (this < 1.0) {
        469.8 * this + 2.391
    } else {
        256.0 * this + 214
    }.toInt()
