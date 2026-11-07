@file:JvmName("MainKt")
package com.scheduleverifier.desktop

import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.Canvas
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.geometry.CornerRadius
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.geometry.Size
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.Path
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.window.Dialog
import androidx.compose.ui.window.DialogProperties
import androidx.compose.ui.window.Window
import androidx.compose.ui.window.WindowPosition
import androidx.compose.ui.window.application
import androidx.compose.ui.window.rememberWindowState
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.pdfbox.Loader
import org.apache.pdfbox.text.PDFTextStripper
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xwpf.usermodel.LineSpacingRule
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.nio.charset.Charset
import java.time.YearMonth
import java.util.Locale
import javax.swing.JFileChooser
import javax.swing.filechooser.FileNameExtensionFilter

// ============================================================================
// 1. ЯВНЫЕ ТИПЫ И МОДЕЛИ ДАННЫХ
// ============================================================================

typealias TrainScheduleMap = Map<String, TrainChangeRow>
typealias DateTrainMap = Map<String, TrainScheduleMap>
typealias StationDateTrainMap = Map<String, DateTrainMap>

enum class MainNavTab(val title: String) {
    UPLOADS("ВЫГРУЗКИ"),
    REPORTS("ОТЧЕТЫ / СВЕРКА"),
    YANDEX("РАСПИСАНИЕ / ЯНДЕКС")
}

enum class DirectionTab(val title: String, val isToMoscow: Boolean) {
    TO_MOSCOW("НА МОСКВУ (НА ИППОДРОМ)", true),
    FROM_MOSCOW("НА ОБЛАСТЬ (ИЗ ИППОДРОМА)", false)
}

enum class TrainFilterType(val label: String) {
    ALL("Все"),
    MCD("МЦД"),
    EXPRESS("Экспрессы / ПДС")
}

enum class AuditStatus(val label: String, val color: Color) {
    OK("СОВПАДАЕТ", Color(0xFF16A34A)),
    CRITICAL("КРИТИЧЕСКАЯ ОШИБКА", Color(0xFFDC2626)),
    TIME_MISMATCH("РАСХОЖДЕНИЕ ВРЕМЕНИ", Color(0xFFD97706)),
    STOP_ERROR("ОШИБКА ОСТАНОВОК", Color(0xFF7C3AED)),
    GLOBAL_INFRA_ERROR("ЗАКРЫТАЯ СТАНЦИЯ", Color(0xFF991B1B))
}

data class StationStopEntry(
    val stationName: String,
    val afishaTime: String = "—",
    val telegramTime: String = "—",
    val exportTime: String = "—",
    val yandexTime: String = "—",
    val isSkippedByTelegram: Boolean = false,
    val isAddedByTelegram: Boolean = false,
    val isTechStop: Boolean = false
)

data class TrainChangeRow(
    val id: String,
    val dateStr: String,
    val station: String,
    val trainNumber: String,
    val isToMoscow: Boolean,
    val isMcd: Boolean,
    val isExpress: Boolean,
    val route: String,
    val telegramBasis: String,
    val afishaTime: String,
    val telegramTime: String,
    val exportTime: String,
    val yandexTime: String,
    val actualChangesText: String,
    val wordExportLine: String,
    val isCanceled: Boolean = false,
    val isAppointed: Boolean = false,
    val skippedStops: List<String> = emptyList(),
    val addedStops: List<String> = emptyList(),
    val globalErrorNote: String? = null,
    val rawTelegramSnippet: String = "",
    val rawExportSnippet: String = "",
    val fullThread: List<StationStopEntry> = emptyList(),
    val auditStatus: AuditStatus = AuditStatus.OK
)

data class AfishaBaseTrain(
    val trainNumber: String,
    val route: String,
    val isToMoscow: Boolean,
    val stationTimes: Map<String, String>
)

data class YandexParsedTrain(
    val trainNumber: String,
    val dateStr: String,
    val station: String,
    val departureTime: String,
    val statusText: String
)

// ============================================================================
// 2. ЦВЕТОВАЯ ПАЛИТРА И КОНСТАНТЫ ПОЛИГОНА
// ============================================================================

object TrainRTheme {
    val NavyDark = Color(0xFF0F172A)
    val SidebarNavy = Color(0xFF0F172A)
    val SidebarActive = Color(0xFF1E293B)
    val RoyalBlue = Color(0xFF1D4ED8)
    val ActionBlue = Color(0xFF2563EB)
    val GoldAccent = Color(0xFFF59E0B)
    val GreenSuccess = Color(0xFF16A34A)
    val RedCancel = Color(0xFFDC2626)
    val AmberWarn = Color(0xFFD97706)
    val PurpleMcd = Color(0xFF7C3AED)
    val YandexRed = Color(0xFFFC3F1D)
    val BorderCol = Color(0xFFE2E8F0)
    val TextDark = Color(0xFF0F172A)
    val TextSub = Color(0xFF64748B)
    val BgLight = Color(0xFFF8FAFC)
}

object PolygonConfig {
    val ALL_STATIONS: List<String> = listOf(
        "Москва", "Останкино", "Лихоборы", "Моссельмаш", "Грачёвская", "Ховрино",
        "Левобережная", "Химки", "Молжаниново", "Новоподрезково", "Подрезково",
        "Сходня", "Фирсановская", "Крюково", "Алабушево", "Радищево",
        "Поваровка", "Поварово-1", "Берёзки-Дачные", "Подсолнечная", "Сенеж",
        "Покровка", "Фроловское", "Стреглово", "Клин", "Ямуга", "Решетниково",
        "Завидово", "Московское Море", "Редкино", "Межево", "Кузьминка",
        "Чуприяновка", "Лазурная", "Тверь", "Конаково ГРЭС"
    )

    val MCD3_STATIONS: Set<String> = setOf(
        "Останкино", "Лихоборы", "Моссельмаш", "Грачёвская", "Грачевская",
        "Ховрино", "Левобережная", "Химки", "Молжаниново", "Новоподрезково",
        "Подрезково", "Сходня", "Фирсановская", "Крюково", "Зеленоград-Крюково"
    )

    val HARD_CLOSED_STATIONS: Set<String> = setOf(
        "Малино", "Рижская", "Петровско-Разумовская", "Лазурная"
    )

    fun normalizeStationName(raw: String): String {
        val clean = raw.trim()
            .replace("Зеленоград-Крюково", "Крюково", ignoreCase = true)
            .replace("Грачевская", "Грачёвская", ignoreCase = true)
            .replace("Березки-Дачные", "Берёзки-Дачные", ignoreCase = true)
            .replace("Москва-Пасс.-Октябрьская", "Москва", ignoreCase = true)
            .replace("Ленинградский вокзал", "Москва", ignoreCase = true)
            .replace("Оп Подрезково", "Подрезково", ignoreCase = true)
        return ALL_STATIONS.firstOrNull { it.equals(clean, ignoreCase = true) } ?: clean
    }

    fun isTrainToMoscow(trainNumber: String, route: String = ""): Boolean {
        if (route.contains("Ипподром", ignoreCase = true) || route.contains("до Москвы", ignoreCase = true) || route.contains("на Москву", ignoreCase = true) || route.contains("47 км", ignoreCase = true)) {
            return true
        }
        if (route.contains("до Крюково", ignoreCase = true) || route.contains("из Москвы", ignoreCase = true) || route.contains("от Москвы", ignoreCase = true) || route.contains("до Твери", ignoreCase = true) || route.contains("до Клина", ignoreCase = true)) {
            return false
        }
        val firstPart = trainNumber.split("/").firstOrNull()?.trim() ?: trainNumber
        val digits = Regex("\\d+").find(firstPart)?.value ?: return true
        val lastDigit = digits.lastOrNull()?.digitToIntOrNull() ?: return true
        return lastDigit % 2 != 0
    }

    fun isMcdTrain(trainNumber: String, route: String): Boolean {
        if (trainNumber.contains("/")) return true
        val digits = Regex("\\d{3,4}").find(trainNumber)?.value?.toIntOrNull() ?: 0
        if (digits in 7200..7499) return true
        if (route.contains("Ипподром", ignoreCase = true) || route.contains("47 км", ignoreCase = true)) return true
        return false
    }

    fun isExpressTrain(trainNumber: String, route: String): Boolean {
        if (isMcdTrain(trainNumber, route)) return false
        val digits = Regex("\\d{3,4}").find(trainNumber)?.value?.toIntOrNull() ?: 0
        return digits in 700..799 || digits in 7000..7999
    }
}

// ============================================================================
// 3. ТОЧКА ВХОДА ПРИЛОЖЕНИЯ
// ============================================================================

fun main() = application {
    val windowState = rememberWindowState(
        width = 1480.dp,
        height = 920.dp,
        position = WindowPosition(Alignment.Center)
    )
    Window(
        onCloseRequest = ::exitApplication,
        state = windowState,
        title = "Train-R • Аналитический портал сверки расписания и генерации отмен"
    ) {
        MaterialTheme {
            TrainRApp()
        }
    }
}

// ============================================================================
// 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ЛОГИКИ
// ============================================================================

fun isFutureStop(currentStation: String, targetStation: String, isToMoscow: Boolean): Boolean {
    val allSt = PolygonConfig.ALL_STATIONS
    val currentIndex = allSt.indexOfFirst { it.equals(currentStation, ignoreCase = true) }
    val targetIndex = allSt.indexOfFirst { it.equals(targetStation, ignoreCase = true) }
    if (currentIndex == -1 || targetIndex == -1) return true
    
    return if (isToMoscow) {
        targetIndex < currentIndex
    } else {
        targetIndex > currentIndex
    }
}

fun extractStationsByRegex(text: String, pattern: String): List<String> {
    val matches = Regex(pattern, RegexOption.IGNORE_CASE).findAll(text)
    val result = mutableListOf<String>()
    for (match in matches) {
        val listStr = match.groupValues[1]
        val parts = listStr.split(Regex(",|\\s+и\\s+", RegexOption.IGNORE_CASE))
        for (p in parts) {
            val norm = PolygonConfig.normalizeStationName(p.replace(".", "").trim())
            if (norm in PolygonConfig.ALL_STATIONS) result.add(norm)
        }
    }
    return result
}

fun buildRemainingStopsNote(
    skippedStops: List<String>,
    addedStops: List<String>,
    isTechStop: Boolean = false,
    isCanceled: Boolean = false,
    pathChanges: String = "",
    forWord: Boolean = false
): String {
    if (isCanceled) return ""
    val parts = mutableListOf<String>()
    
    val distinctSkipped = skippedStops.distinct()
    if (distinctSkipped.isNotEmpty()) {
        parts.add("отменены остановки: " + distinctSkipped.joinToString(", "))
    }
    
    val distinctAdded = addedStops.distinct()
    if (distinctAdded.isNotEmpty()) {
        parts.add("назначены остановки: " + distinctAdded.joinToString(", "))
    }
    
    if (isTechStop) {
        parts.add("тех. стоянка (без посадки/высадки)")
    }
    
    val cleanPath = pathChanges.replace(Regex("(?i)(?:^|[^а-яА-ЯёЁ])Изменений\\s*нет\\.?"), "").trim()
    if (cleanPath.isNotEmpty()) {
        parts.add(cleanPath)
    }
    
    if (parts.isEmpty()) {
        return if (forWord) "" else "по графику со всеми остановками"
    }
    return parts.joinToString(". ").replace(Regex("\\.{2,}"), ".")
}

fun calculateMinutesDiff(timeA: String, timeB: String): String {
    val reg = Regex("(\\d{1,2}):(\\d{2})")
    val mA = reg.find(timeA) ?: return "—"
    val mB = reg.find(timeB) ?: return "—"
    val minA = mA.groupValues[1].toInt() * 60 + mA.groupValues[2].toInt()
    val minB = mB.groupValues[1].toInt() * 60 + mB.groupValues[2].toInt()
    val diff = minB - minA
    return when {
        diff == 0 -> "0 мин"
        diff > 0 -> "+$diff мин"
        else -> "$diff мин"
    }
}

fun parseTrainTimeForSorting(timeStr: String): Int {
    if (timeStr == "—" || timeStr.contains("ОТМЕН", true)) return 9999
    val reg = Regex("(\\d{1,2}):(\\d{2})")
    val m = reg.find(timeStr) ?: return 9999
    val hours = m.groupValues[1].toInt()
    val minutes = m.groupValues[2].toInt()
    
    val adjustedHours = if (hours in 0..3) hours + 24 else hours
    return adjustedHours * 60 + minutes
}

fun getBaseTimeForSorting(row: TrainChangeRow): String {
    val regex = Regex("(?i)отправлением\\s+в\\s+(\\d{1,2}:\\d{2})")
    val match = regex.find(row.wordExportLine) ?: regex.find(row.actualChangesText)
    
    if (match != null) {
        return match.groupValues[1].padStart(5, '0')
    }
    
    if (row.afishaTime != "—") return row.afishaTime
    if (row.exportTime != "—" && !row.exportTime.contains("ОТМЕН", true)) return row.exportTime.take(5)
    if (row.telegramTime != "—" && !row.telegramTime.contains("ОТМЕН", true)) return row.telegramTime.take(5)
    return "99:99"
}

fun getSortingTime(row: TrainChangeRow): Int {
    val actualRegex = Regex("(?i)отправится\\s+в\\s+(\\d{1,2}:\\d{2})")
    val actualMatch = actualRegex.find(row.wordExportLine) ?: actualRegex.find(row.actualChangesText)
    if (actualMatch != null) {
        return parseTrainTimeForSorting(actualMatch.groupValues[1])
    }
    
    val baseRegex = Regex("(?i)отправлением\\s+в\\s+(\\d{1,2}:\\d{2})")
    val baseMatch = baseRegex.find(row.wordExportLine) ?: baseRegex.find(row.actualChangesText)
    if (baseMatch != null) {
        return parseTrainTimeForSorting(baseMatch.groupValues[1])
    }
    
    return parseTrainTimeForSorting(row.afishaTime)
}

fun mergeSameTrainRowsForDay(rows: List<TrainChangeRow>): List<TrainChangeRow> {
    if (rows.size <= 1) return rows
    val grouped = rows.groupBy { "${it.dateStr}|${it.station}|${it.trainNumber}" }
    val mergedList = mutableListOf<TrainChangeRow>()

    for ((_, sameTrains) in grouped) {
        if (sameTrains.size == 1) {
            mergedList.add(sameTrains.first())
        } else {
            val base = sameTrains.first()
            val combinedSkipped = sameTrains.flatMap { it.skippedStops }.distinct()
            val combinedAdded = sameTrains.flatMap { it.addedStops }.distinct()
            val combinedCanceled = sameTrains.any { it.isCanceled }
            val combinedAppointed = sameTrains.any { it.isAppointed }
            
            val finalSkipped = if (combinedCanceled) emptyList() else combinedSkipped

            val latestTelegramTime = sameTrains.map { it.telegramTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: base.telegramTime
            val latestExportTime = sameTrains.map { it.exportTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: base.exportTime
            val latestYandexTime = sameTrains.map { it.yandexTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: base.yandexTime

            val mergedActualText = sameTrains.map { it.actualChangesText.trim() }
                .filter { it.isNotEmpty() }
                .maxByOrNull { it.length } ?: ""
            
            val mergedRawTg = sameTrains.map { it.rawTelegramSnippet }.filter { it.isNotBlank() }.distinct().joinToString("\n---\n")
            val mergedRawExp = sameTrains.map { it.rawExportSnippet }.filter { it.isNotBlank() }.distinct().joinToString("\n---\n")
            
            val exportRow = sameTrains.find { it.rawExportSnippet.isNotBlank() }
            val mergedWordLine = exportRow?.wordExportLine?.takeIf { it.isNotBlank() } 
                ?: sameTrains.map { it.wordExportLine.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length } ?: ""

            val mergedBasis = sameTrains.map { it.telegramBasis.trim() }
                .filter { it.isNotEmpty() && it != "—" }
                .distinct()
                .joinToString(", ")
                .ifEmpty { base.telegramBasis }

            val mergedError = sameTrains.mapNotNull { it.globalErrorNote }.distinct().joinToString("; ").ifEmpty { null }

            val evaluatedStatus = evaluateFourWayAuditStatus(
                afishaTime = base.afishaTime,
                telegramTime = latestTelegramTime,
                exportTime = latestExportTime,
                yandexTime = latestYandexTime,
                isCanceled = combinedCanceled,
                skippedStops = finalSkipped,
                globalError = mergedError
            )

            mergedList.add(
                base.copy(
                    telegramBasis = mergedBasis,
                    telegramTime = latestTelegramTime,
                    exportTime = latestExportTime,
                    yandexTime = latestYandexTime,
                    actualChangesText = mergedActualText,
                    wordExportLine = mergedWordLine,
                    isCanceled = combinedCanceled,
                    isAppointed = combinedAppointed,
                    skippedStops = finalSkipped,
                    addedStops = combinedAdded,
                    globalErrorNote = mergedError,
                    rawTelegramSnippet = mergedRawTg,
                    rawExportSnippet = mergedRawExp,
                    auditStatus = evaluatedStatus
                )
            )
        }
    }
    return mergedList.sortedBy { getSortingTime(it) }
}

fun evaluateFourWayAuditStatus(
    afishaTime: String,
    telegramTime: String,
    exportTime: String,
    yandexTime: String,
    isCanceled: Boolean,
    skippedStops: List<String>,
    globalError: String?
): AuditStatus {
    if (!globalError.isNullOrBlank()) return AuditStatus.GLOBAL_INFRA_ERROR
    if (isCanceled && yandexTime != "—" && !yandexTime.contains("ОТМЕН", true)) {
        return AuditStatus.CRITICAL
    }
    if (telegramTime != "—" && exportTime != "—" && telegramTime != exportTime && !isCanceled) {
        return AuditStatus.TIME_MISMATCH
    }
    if (telegramTime != "—" && yandexTime != "—" && telegramTime != yandexTime && !isCanceled) {
        return AuditStatus.TIME_MISMATCH
    }
    if (skippedStops.isNotEmpty() && yandexTime != "—" && yandexTime == afishaTime) {
        return AuditStatus.STOP_ERROR
    }
    return AuditStatus.OK
}

// ============================================================================
// 5. ПАРСИНГ ФАЙЛОВ
// ============================================================================

object FileParsers {
    private val cacheDir = File(System.getProperty("user.home"), ".train_r_system")
    val cachedAfishaFile = File(cacheDir, "cached_afisha.xlsx")

    fun ensureCacheAfisha(sourceFile: File): File {
        if (!cacheDir.exists()) cacheDir.mkdirs()
        sourceFile.copyTo(cachedAfishaFile, overwrite = true)
        return cachedAfishaFile
    }

    fun readSmartTextFile(file: File): String {
        val ext = file.extension.lowercase(Locale.getDefault())
        if (ext == "docx") {
            FileInputStream(file).use { fis ->
                XWPFDocument(fis).use { doc ->
                    return doc.paragraphs.joinToString("\n") { it.text }
                }
            }
        }
        if (ext == "pdf") {
            Loader.loadPDF(file).use { pdDoc ->
                val stripper = PDFTextStripper()
                return stripper.getText(pdDoc)
            }
        }
        val bytes = file.readBytes()
        // ИСПРАВЛЕНИЕ: Удаляем скрытые BOM символы, из-за которых ломался парсинг первой станции (Конаково ГРЭС)
        val utf8Str = String(bytes, Charsets.UTF_8).replace("\uFEFF", "")
        val cp1251Str = String(bytes, Charset.forName("windows-1251")).replace("\uFEFF", "")
        val cyrillicCountUtf8 = Regex("[А-Яа-яЁё]").findAll(utf8Str).count()
        val cyrillicCount1251 = Regex("[А-Яа-яЁё]").findAll(cp1251Str).count()

        return if ( utf8Str.contains("\uFFFD") || cyrillicCount1251 > cyrillicCountUtf8 ) cp1251Str else utf8Str
    }

    fun parseAfishaExcel(file: File): Map<String, AfishaBaseTrain> {
        val result = mutableMapOf<String, AfishaBaseTrain>()
        if (!file.exists()) return result
        try {
            FileInputStream(file).use { fis ->
                WorkbookFactory.create(fis).use { wb ->
                    for (sIdx in 0 until wb.numberOfSheets) {
                        val sheet = wb.getSheetAt(sIdx) ?: continue
                        var trainHeaderRowIdx = -1
                        val colToTrainNum = mutableMapOf<Int, String>()
                        for (rIdx in 0..minOf(15, sheet.lastRowNum)) {
                            val row = sheet.getRow(rIdx) ?: continue
                            for (cIdx in 0 until row.lastCellNum.coerceAtLeast(0)) {
                                val cellVal = extractCellText(row.getCell(cIdx))
                                val m = Regex("\\b([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)\\b").find(cellVal)
                                if (m != null) {
                                    colToTrainNum[cIdx] = m.groupValues[1]
                                    trainHeaderRowIdx = rIdx
                                }
                            }
                            if (colToTrainNum.size >= 3) break
                        }
                        if (trainHeaderRowIdx != -1 && colToTrainNum.isNotEmpty()) {
                            val trainStationTimes = mutableMapOf<String, MutableMap<String, String>>()
                            for (rIdx in (trainHeaderRowIdx + 1)..sheet.lastRowNum) {
                                val row = sheet.getRow(rIdx) ?: continue
                                val rawStation = extractCellText(row.getCell(0))
                                val station = PolygonConfig.normalizeStationName(rawStation)
                                if (station !in PolygonConfig.ALL_STATIONS) continue
                                for ((cIdx, trNum) in colToTrainNum) {
                                    val timeStr = extractCellText(row.getCell(cIdx))
                                    val cleanTime = Regex("(\\d{1,2}:\\d{2})").find(timeStr)?.groupValues?.get(1)?.padStart(5, '0')
                                    if (cleanTime != null) trainStationTimes.getOrPut(trNum) { mutableMapOf() }[station] = cleanTime
                                }
                            }
                            for ((trNum, stMap) in trainStationTimes) {
                                val isToMsk = PolygonConfig.isTrainToMoscow(trNum)
                                val routeStr = if (isToMsk) "Крюково/Тверь — Москва" else "Москва — Крюково/Тверь"
                                result[trNum] = AfishaBaseTrain(
                                    trainNumber = trNum, route = routeStr, isToMoscow = isToMsk, stationTimes = stMap
                                )
                            }
                        }
                    }
                }
            }
        } catch (_: Exception) {}
        return result
    }

    private fun extractCellText(cell: org.apache.poi.ss.usermodel.Cell?): String {
        if (cell == null) return ""
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue.trim()
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    val date = cell.localDateTimeCellValue
                    String.format("%02d:%02d", date.hour, date.minute)
                } else cell.numericCellValue.toLong().toString()
            }
            CellType.FORMULA -> try { cell.stringCellValue.trim() } catch (_: Exception) { cell.numericCellValue.toInt().toString() }
            else -> ""
        }
    }

    fun parseTelegramFiles(files: List<File>, afishaMap: Map<String, AfishaBaseTrain>): StationDateTrainMap {
        val rawRowsByStationDate = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        val trainRegex = Regex("(?i)(?:нр\\.|№)\\s*([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)")
        val dateRegex = Regex("(?i)(\\d{1,2}(?:\\s*,\\s*\\d{1,2})*)\\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)")
        
        val timePairRegex = Regex("([А-Яа-яЁё\\-\\s\\d]+?)\\s+(\\d{1,2}[.:]\\d{2}(?::\\d{2})?)(?:\\s*[-–—]\\s*(\\d{1,2}[.:]\\d{2}(?::\\d{2})?))?")

        for (file in files) {
            val fullText = readSmartTextFile(file)
            val isClarification = fullText.contains("Внести разъяснение", ignoreCase = true)
            val basisStr = Regex("(?i)(?:телеграмм[аы]|ТГ|№)\\s*([\\d/\\-А-Яа-я]+)").find(fullText)?.value ?: file.nameWithoutExtension

            val paragraphs = mutableListOf<String>()
            var currentPara = ""
            for (line in fullText.lines()) {
                val trimmed = line.trim()
                if (trimmed.isEmpty()) continue
                
                val isNewTrainLine = Regex("^(?i)(?:нр\\.|№)\\s*[1-7]\\d{2,3}").containsMatchIn(trimmed)
                val isMajorHeader = Regex("^(?i)(?:\\d+\\.|Отменить|Отменяются|Изменить)").containsMatchIn(trimmed) && !Regex("(?i)(?:нр\\.|№)").containsMatchIn(trimmed)
                
                if (isNewTrainLine || isMajorHeader) {
                    if (currentPara.isNotEmpty()) paragraphs.add(currentPara)
                    currentPara = trimmed
                } else {
                    if (currentPara.isNotEmpty()) currentPara += " $trimmed"
                    else currentPara = trimmed
                }
            }
            if (currentPara.isNotEmpty()) paragraphs.add(currentPara)

            var isGlobalCancelMode = false
            var currentDates = mutableListOf("19 октября")

            for (para in paragraphs) {
                val cleanPara = para.replace("\n", " ").replace("\r", "").trim()
                if (cleanPara.isEmpty()) continue

                val extractedDates = dateRegex.findAll(cleanPara).map { it.groupValues[1] + " " + it.groupValues[2].lowercase(Locale.getDefault()) }.toList()
                if (extractedDates.isNotEmpty()) currentDates = extractedDates.toMutableList()

                // ИСПРАВЛЕНИЕ: Использование RegexOption.IGNORE_CASE для кириллицы
                if (Regex("(?:^|[^а-яА-ЯёЁ])(отменить|отменяются)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)) {
                    isGlobalCancelMode = true
                } else if (Regex("(?:^|[^а-яА-ЯёЁ])(изменить|назначить)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)) {
                    isGlobalCancelMode = false
                }

                val trainMatches = trainRegex.findAll(cleanPara).map { it.groupValues[1] }.distinct().toList()
                if (trainMatches.isEmpty()) continue

                val localCancel = Regex("(?:^|[^а-яА-ЯёЁ])(отменяется|отмен[её]н|отменить|отмена)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara) && 
                                  !Regex("отменен[а-я]*\\s*остановк", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara) &&
                                  !Regex("без\\s+остановк", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)
                
                val hasNewTime = Regex("отправится\\s+в", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)
                val hasStops = Regex("(?:отменен[ыа-я]*|назначен[ыа-я]*|без)\\s*останов[ок]", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)
                val isImplicitCancel = !hasNewTime && !hasStops && !Regex("назначается", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)

                val isCanceled = isGlobalCancelMode || localCancel || isImplicitCancel
                val isAppointed = Regex("назначается|назначить", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)

                val closedStationError = PolygonConfig.HARD_CLOSED_STATIONS.firstOrNull { 
                    cleanPara.contains(it, true) && !cleanPara.contains("без остановки", true) 
                }?.let { "ОШИБКА ТГ: Остановка на закрытой $it!" }

                val pathRegex = Regex("(?i)(\\([^)]*пут[иь][^)]*\\)|по\\s*\\d+\\s*пут[иь]|проследует\\s*по\\s*\\d+\\s*пут[иь][^.,;()]*)")
                val pathStr = pathRegex.findAll(cleanPara).map { it.value }.joinToString(" ")

                val skippedStops = mutableListOf<String>()
                val addedStops = mutableListOf<String>()
                
                skippedStops.addAll(extractStationsByRegex(cleanPara, "(?i)отменен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)"))
                skippedStops.addAll(extractStationsByRegex(cleanPara, "(?i)без\\s*останов[ок]+\\s*:?\\s*([^.;(]+)"))
                addedStops.addAll(extractStationsByRegex(cleanPara, "(?i)назначен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)"))

                val stationDepTimes = mutableMapOf<String, String>()
                val stopMatches = timePairRegex.findAll(cleanPara).toList()
                
                for ((idx, sm) in stopMatches.withIndex()) {
                    val rawSt = sm.groupValues[1].trim()
                    val stName = PolygonConfig.normalizeStationName(rawSt)
                    if (stName !in PolygonConfig.ALL_STATIONS) continue

                    val t1 = sm.groupValues[2].replace('.', ':').take(5).padStart(5, '0')
                    val t2 = sm.groupValues[3].replace('.', ':').take(5).let { if (it.isNotEmpty()) it.padStart(5, '0') else "" }
                    val isTech = cleanPara.contains("тех. стоянка", true) || cleanPara.contains("стоянка техническая", true) || cleanPara.contains("без пассажиров", true)

                    if (isClarification) {
                        if (cleanPara.contains("посадка", true)) addedStops.add(stName)
                        else if (cleanPara.contains("техническая", true)) skippedStops.add(stName)
                    } else {
                        if (t2.isNotEmpty()) {
                            if (isTech) skippedStops.add(stName) else stationDepTimes[stName] = t2
                        } else {
                            if (idx == 0 || idx == stopMatches.lastIndex) {
                                stationDepTimes[stName] = t1 
                            } else {
                                skippedStops.add(stName)
                            }
                        }
                    }
                }

                val mentionedStations = PolygonConfig.ALL_STATIONS.filter { cleanPara.contains(it, true) }
                val targetStations = when {
                    stationDepTimes.isNotEmpty() -> stationDepTimes.keys
                    mentionedStations.isNotEmpty() -> mentionedStations.toSet()
                    isCanceled || isAppointed -> setOf("Крюково", "Химки", "Сходня", "Клин", "Подсолнечная", "Тверь")
                    else -> emptySet()
                }

                if (targetStations.isEmpty()) continue
                
                val validSkippedStops = skippedStops.filter { it !in PolygonConfig.HARD_CLOSED_STATIONS }
                val validAddedStops = addedStops.filter { it !in PolygonConfig.HARD_CLOSED_STATIONS }

                for (trainNum in trainMatches) {
                    for (st in targetStations) {
                        val isToMsk = PolygonConfig.isTrainToMoscow(trainNum, cleanPara)
                        if (st == "Левобережная" && !isToMsk) continue

                        val timesFound = Regex("(\\d{1,2}:\\d{2})").findAll(cleanPara).map { it.groupValues[1].padStart(5, '0') }.toList()
                        val baseTimeMatch = Regex("(?i)отправлением\\s+в\\s+(\\d{1,2}:\\d{2})").find(cleanPara)
                        val extractedBaseTime = baseTimeMatch?.groupValues?.get(1)?.padStart(5, '0') ?: timesFound.firstOrNull() ?: "—"
                        
                        val afishaT = afishaMap[trainNum]?.stationTimes?.get(st) ?: extractedBaseTime
                        
                        val isSkippedHere = st in skippedStops
                        val isCanceledAtThisStation = isCanceled || isSkippedHere

                        val rawTgT = stationDepTimes[st] ?: afishaT
                        val tgT = if (isCanceledAtThisStation) {
                            val t = if (afishaT != "—") afishaT else if (rawTgT != "—") rawTgT else ""
                            if (t.isNotEmpty()) "$t ОТМЕНА" else "ОТМЕНА"
                        } else {
                            rawTgT
                        }
                        
                        val routeStr = when {
                            cleanPara.contains("Ипподром", true) || cleanPara.contains("47 км", true) -> if (isToMsk) "до Ипподрома (47 км)" else "до Крюково"
                            cleanPara.contains("Крюково", true) && !isToMsk -> "до Крюково"
                            cleanPara.contains("Тверь", true) && !isToMsk -> "до Твери"
                            else -> if (isToMsk) "до Москвы" else "из Москвы"
                        }
                        
                        val isMcd = PolygonConfig.isMcdTrain(trainNum, cleanPara)
                        val isExp = PolygonConfig.isExpressTrain(trainNum, cleanPara)
                        
                        // ИСПРАВЛЕНИЕ: Для телеграмм мы честно фильтруем по пройденным станциям
                        val finalSkippedForNote = validSkippedStops.filter { it != st && isFutureStop(st, it, isToMsk) }
                        val finalAddedForNote = validAddedStops.filter { it != st && isFutureStop(st, it, isToMsk) }
                        
                        val remNoteUI = buildRemainingStopsNote(finalSkippedForNote, finalAddedForNote, false, isCanceledAtThisStation, pathStr, forWord = false)
                        val remNoteWord = buildRemainingStopsNote(finalSkippedForNote, finalAddedForNote, false, isCanceledAtThisStation, pathStr, forWord = true)
                        
                        val changeSummary = when {
                            isCanceledAtThisStation -> "Поезд №$trainNum отменён по станции $st"
                            isAppointed -> "Назначается поезд №$trainNum" + (if (tgT != "—") " отпр. в $tgT" else "") + (if (remNoteUI.isNotEmpty()) " ($remNoteUI)" else "")
                            else -> "Поезд №$trainNum" + (if (tgT != "—") " отпр. в $tgT" else "") + (if (remNoteUI.isNotEmpty()) " ($remNoteUI)" else "")
                        }

                        val wordText = if (isCanceledAtThisStation) {
                            if (afishaT != "—") "№$trainNum $routeStr отправлением в $afishaT - ОТМЕНЕН"
                            else "№$trainNum $routeStr - ОТМЕНЕН"
                        } else {
                            val base = if (afishaT != "—" && rawTgT != "—") {
                                if (afishaT == rawTgT) "№$trainNum $routeStr отправлением в $afishaT"
                                else "№$trainNum $routeStr отправлением в $afishaT - отправится в $rawTgT"
                            } else if (afishaT != "—") {
                                "№$trainNum $routeStr отправлением в $afishaT"
                            } else {
                                "№$trainNum $routeStr"
                            }
                            if (remNoteWord.isNotEmpty()) "$base. $remNoteWord".replace(" .", ".").trim() else base
                        }

                        for (dateKey in currentDates.distinct()) {
                            val rowObj = TrainChangeRow(
                                id = "${dateKey}_${st}_${trainNum}_${System.nanoTime()}",
                                dateStr = dateKey,
                                station = st,
                                trainNumber = trainNum,
                                isToMoscow = isToMsk,
                                isMcd = isMcd,
                                isExpress = isExp,
                                route = routeStr,
                                telegramBasis = basisStr,
                                afishaTime = afishaT,
                                telegramTime = tgT,
                                exportTime = tgT,
                                yandexTime = "—",
                                actualChangesText = changeSummary.trim(),
                                wordExportLine = wordText,
                                isCanceled = isCanceledAtThisStation,
                                isAppointed = isAppointed,
                                skippedStops = skippedStops,
                                addedStops = addedStops,
                                globalErrorNote = closedStationError,
                                rawTelegramSnippet = cleanPara,
                                rawExportSnippet = "",
                                auditStatus = evaluateFourWayAuditStatus(afishaT, tgT, tgT, "—", isCanceledAtThisStation, skippedStops, closedStationError)
                            )
                            rawRowsByStationDate.getOrPut(st) { mutableMapOf() }.getOrPut(dateKey) { mutableListOf() }.add(rowObj)
                        }
                    }
                }
            }
        }

        val finalMap = mutableMapOf<String, Map<String, Map<String, TrainChangeRow>>>()
        for ((st, dateMap) in rawRowsByStationDate) {
            val innerDateMap = mutableMapOf<String, Map<String, TrainChangeRow>>()
            for ((dt, list) in dateMap) {
                innerDateMap[dt] = mergeSameTrainRowsForDay(list).associateBy { it.trainNumber }
            }
            finalMap[st] = innerDateMap
        }
        return finalMap
    }

    fun parseSummaryExportFile(
        file: File,
        afishaMap: Map<String, AfishaBaseTrain>,
        existingMap: StationDateTrainMap
    ): StationDateTrainMap {
        val text = readSmartTextFile(file)
        val mutableResult = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        for ((st, dtMap) in existingMap) {
            val mDt = mutableMapOf<String, MutableList<TrainChangeRow>>()
            for ((dt, trMap) in dtMap) mDt[dt] = trMap.values.toMutableList()
            mutableResult[st] = mDt
        }

        // ИСПРАВЛЕНИЕ: Очищаем стартовую станцию по умолчанию, чтобы левые поезда не попадали в Крюково.
        var currentStation = ""
        var currentDate = "19 октября"
        
        val stationHeaderRegex = Regex("^\\s*([А-Яа-яЁё\\-\\s\\d.]+?)\\s+Изменения\\s+в\\s+графике", RegexOption.IGNORE_CASE)
        val standaloneStationRegex = Regex("^\\s*(?:Станция|ст\\.)?\\s*([А-Яа-яЁё\\-\\s\\d.]+)\\s*$", RegexOption.IGNORE_CASE)
        val dateLineRegex = Regex("^\\s*(\\d{1,2}(?:\\s*,\\s*\\d{1,2})*)\\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*$", RegexOption.IGNORE_CASE)
        val trainNumberRegex = Regex("№\\s*([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)")

        val paragraphs = mutableListOf<String>()
        var currentPara = ""
        for (line in text.lines()) {
            val trimmed = line.trim()
            if (trimmed.isEmpty()) continue
            if (Regex("(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)", RegexOption.IGNORE_CASE).containsMatchIn(trimmed)) continue

            val stHeaderMatch = stationHeaderRegex.find(trimmed)
            if (stHeaderMatch != null && PolygonConfig.normalizeStationName(stHeaderMatch.groupValues[1]) in PolygonConfig.ALL_STATIONS) {
                if (currentPara.isNotEmpty()) paragraphs.add(currentPara)
                paragraphs.add(trimmed)
                currentPara = ""
                continue
            }
            
            val stSoloMatch = standaloneStationRegex.find(trimmed)
            if (stSoloMatch != null && PolygonConfig.normalizeStationName(stSoloMatch.groupValues[1]) in PolygonConfig.ALL_STATIONS) {
                if (currentPara.isNotEmpty()) paragraphs.add(currentPara)
                paragraphs.add(trimmed)
                currentPara = ""
                continue
            }

            if (dateLineRegex.matches(trimmed)) {
                if (currentPara.isNotEmpty()) paragraphs.add(currentPara)
                paragraphs.add(trimmed)
                currentPara = ""
                continue
            }
            
            if (trainNumberRegex.containsMatchIn(trimmed)) {
                if (currentPara.isNotEmpty()) paragraphs.add(currentPara)
                currentPara = trimmed
            } else {
                if (currentPara.isNotEmpty()) currentPara += " $trimmed" else currentPara = trimmed
            }
        }
        if (currentPara.isNotEmpty()) paragraphs.add(currentPara)

        var isGlobalCancelMode = false

        for (paragraph in paragraphs) {
            var cleanParagraph = paragraph.replace(Regex("[.\\s]*Изменений\\s*нет\\.?", RegexOption.IGNORE_CASE), "")
                .replace(Regex("Бологое[- ]*Моск(?:овское)?\\.?", RegexOption.IGNORE_CASE), "")
                .replace(Regex("\\.{2,}"), ".")
                .replace(Regex("\\s{2,}"), " ")
                .trim()

            if (cleanParagraph.isEmpty()) continue

            if (Regex("(?:^|[^а-яА-ЯёЁ])(отменить|отменяются)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)) {
                isGlobalCancelMode = true
            } else if (Regex("(?:^|[^а-яА-ЯёЁ])(изменить|назначить)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)) {
                isGlobalCancelMode = false
            }

            val stHeaderMatch = stationHeaderRegex.find(cleanParagraph)
            if (stHeaderMatch != null) {
                PolygonConfig.normalizeStationName(stHeaderMatch.groupValues[1]).takeIf { it in PolygonConfig.ALL_STATIONS }?.let { currentStation = it }
                continue
            }
            val stSoloMatch = standaloneStationRegex.find(cleanParagraph)
            if (stSoloMatch != null) {
                PolygonConfig.normalizeStationName(stSoloMatch.groupValues[1]).takeIf { it in PolygonConfig.ALL_STATIONS }?.let { currentStation = it }
                continue
            }
            val dtMatch = dateLineRegex.find(cleanParagraph)
            if (dtMatch != null) {
                val daysList = dtMatch.groupValues[1].split(",").map { it.trim() }.filter { it.isNotEmpty() }
                val monthName = dtMatch.groupValues[2].lowercase(Locale.getDefault())
                if (daysList.isNotEmpty()) currentDate = "${daysList.first()} $monthName"
                continue
            }

            if (currentStation.isEmpty()) continue

            val trainMatches = trainNumberRegex.findAll(cleanParagraph).map { it.groupValues[1] }.distinct().toList()
            if (trainMatches.isEmpty()) continue

            val localCancel = Regex("(?:^|[^а-яА-ЯёЁ])(отменяется|отмен[её]н|отменить|отмена)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph) && 
                              !Regex("отменен[а-я]*\\s+остановк", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph) &&
                              !Regex("без\\s+остановк", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)
            
            val hasNewTime = Regex("отправится\\s+в", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)
            val hasStops = Regex("(?:отменен[ыа-я]*|назначен[ыа-я]*|без)\\s*останов[ок]", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)
            val isImplicitCancel = !hasNewTime && !hasStops && !Regex("назначается", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)

            val isCanceled = isGlobalCancelMode || localCancel || isImplicitCancel
            val isAppointed = Regex("назначается", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)
            
            cleanParagraph = cleanParagraph.replace(Regex("\\s*\\(приб\\.?\\s*\\d{1,2}:\\d{2}\\)", RegexOption.IGNORE_CASE), "").trim()

            val skippedStops = mutableListOf<String>()
            skippedStops.addAll(extractStationsByRegex(cleanParagraph, "(?i)отменен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)"))
            skippedStops.addAll(extractStationsByRegex(cleanParagraph, "(?i)без\\s*останов[ок]+\\s*:?\\s*([^.;(]+)"))
            
            val addedStops = extractStationsByRegex(cleanParagraph, "(?i)назначен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)")

            for (trNum in trainMatches) {
                val isToMsk = PolygonConfig.isTrainToMoscow(trNum, cleanParagraph)
                if (currentStation == "Левобережная" && !isToMsk) continue

                val isSkippedHere = currentStation in skippedStops
                val isCanceledAtThisStation = isCanceled || isSkippedHere

                val timesFound = Regex("(\\d{1,2}:\\d{2})").findAll(cleanParagraph).map { it.groupValues[1].padStart(5, '0') }.toList()
                val baseTimeMatch = Regex("(?i)отправлением\\s+в\\s+(\\d{1,2}:\\d{2})").find(cleanParagraph)
                val extractedBaseTime = baseTimeMatch?.groupValues?.get(1)?.padStart(5, '0') ?: timesFound.firstOrNull() ?: "—"
                val afishaT = afishaMap[trNum]?.stationTimes?.get(currentStation) ?: extractedBaseTime
                
                val rawExpT = if (timesFound.size >= 2) timesFound.last() else if (timesFound.size == 1) timesFound.first() else afishaT
                
                val expT = if (isCanceledAtThisStation) {
                    val t = if (afishaT != "—") afishaT else if (rawExpT != "—") rawExpT else ""
                    if (t.isNotEmpty()) "$t ОТМЕНА" else "ОТМЕНА"
                } else {
                    rawExpT
                }
                
                // ИСПРАВЛЕНИЕ: Выгрузка переносится 100% ДОСЛОВНО (ничего не обрезаем).
                // Но если в самой фразе диспетчер забыл написать "ОТМЕНЕН", а по факту это отмена (нет нового времени),
                // программа сама дописывает "- ОТМЕНЕН", чтобы поезд корректно ушел в отчет.
                val wordText = if (isCanceledAtThisStation && !Regex("отмен", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph)) {
                    "$cleanParagraph - ОТМЕНЕН"
                } else {
                    cleanParagraph
                }

                val row = TrainChangeRow(
                    id = "${currentDate}_${currentStation}_${trNum}_${System.nanoTime()}",
                    dateStr = currentDate, station = currentStation, trainNumber = trNum, isToMoscow = isToMsk,
                    isMcd = PolygonConfig.isMcdTrain(trNum, cleanParagraph), isExpress = PolygonConfig.isExpressTrain(trNum, cleanParagraph), route = "Выгрузка маршрут",
                    telegramBasis = "Выгрузка (${file.name})", afishaTime = afishaT, telegramTime = expT, exportTime = expT, yandexTime = "—",
                    actualChangesText = cleanParagraph, wordExportLine = wordText, isCanceled = isCanceledAtThisStation, isAppointed = isAppointed, rawTelegramSnippet = "", rawExportSnippet = cleanParagraph
                )
                mutableResult.getOrPut(currentStation) { mutableMapOf() }.getOrPut(currentDate) { mutableListOf() }.add(row)
            }
        }

        val out = mutableMapOf<String, Map<String, Map<String, TrainChangeRow>>>()
        for ((st, dtMap) in mutableResult) {
            val dOut = mutableMapOf<String, Map<String, TrainChangeRow>>()
            for ((dt, rows) in dtMap) dOut[dt] = mergeSameTrainRowsForDay(rows).associateBy { it.trainNumber }
            out[st] = dOut
        }
        return out
    }

    fun parseYandexPdf(file: File): List<YandexParsedTrain> {
        val text = readSmartTextFile(file)
        val parsed = mutableListOf<YandexParsedTrain>()
        val lineRegex = Regex("([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?).*?(\\d{1,2}:\\d{2})")
        for (line in text.lines()) {
            val m = lineRegex.find(line) ?: continue
            val trNum = m.groupValues[1]
            val depTime = m.groupValues[2].padStart(5, '0')
            parsed.add(YandexParsedTrain(trNum, "19 октября", "Крюково", depTime, if (Regex("отмен", RegexOption.IGNORE_CASE).containsMatchIn(line)) "Отменён" else "По расписанию"))
        }
        return parsed
    }
}

// ============================================================================
// 6. ГЕНЕРАТОР ФАЙЛОВ WORD (.DOCX)
// ============================================================================

object WordCancelGenerator {
    private val arrivalExcludeRegex = Regex("(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)", RegexOption.IGNORE_CASE)
    private val nightNoteRegex = Regex("\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)", RegexOption.IGNORE_CASE)

    fun getDirectionSubtitle(station: String, isToMoscow: Boolean): String = if (isToMoscow) {
        if (station in PolygonConfig.MCD3_STATIONS) "на Ленинградский вокзал, Ипподром" else "на Москву"
    } else "от Москвы"

    fun generateStationWordFiles(selectedStations: Set<String>, selectedDates: Set<String>, directionTab: DirectionTab, allData: StationDateTrainMap): Pair<Int, File> {
        val desktopDir = File(System.getProperty("user.home"), "Desktop")
        val targetFolder = File(desktopDir, "Отмены по станциям Word")
        if (!targetFolder.exists()) targetFolder.mkdirs()

        var generatedCount = 0
        val isToMoscow = directionTab.isToMoscow

        for (station in selectedStations) {
            if (station == "Левобережная" && !isToMoscow) continue
            val stationDateMap = allData[station] ?: continue
            val datesToProcess = if (selectedDates.isEmpty()) stationDateMap.keys else stationDateMap.keys.intersect(selectedDates)
            if (datesToProcess.isEmpty()) continue

            val signatureToDates = linkedMapOf<String, MutableList<String>>()
            val dateToCleanLines = mutableMapOf<String, List<String>>()

            for (dateStr in datesToProcess) {
                val trainMap = stationDateMap[dateStr] ?: continue
                val filteredRows = trainMap.values
                    .filter { it.isToMoscow == isToMoscow && !arrivalExcludeRegex.containsMatchIn(it.wordExportLine) && !arrivalExcludeRegex.containsMatchIn(it.actualChangesText) }
                if (filteredRows.isEmpty()) continue

                val sortedLines = filteredRows.sortedBy { row ->
                    getSortingTime(row)
                }.map { it.wordExportLine.trim() }.distinct()

                dateToCleanLines[dateStr] = sortedLines
                val normalizedSignature = sortedLines.joinToString("\n") { nightNoteRegex.replace(it, "").trim() }
                signatureToDates.getOrPut(normalizedSignature) { mutableListOf() }.add(dateStr)
            }

            if (signatureToDates.isEmpty()) continue
            val fileNameSuffix = if (isToMoscow) "на Москву" else "от Москвы"
            val outFile = File(targetFolder, "$station $fileNameSuffix.docx")

            XWPFDocument().use { doc ->
                val subtitle = getDirectionSubtitle(station, isToMoscow)
                addCenteredBoldHeader(doc, "$station  Изменения в графике движения эл/поездов")
                addCenteredBoldHeader(doc, subtitle)
                var currentPageLineCount = 2

                for ((_, groupedDates) in signatureToDates) {
                    val isMergedMultipleDays = groupedDates.size > 1
                    val formattedDatesHeader = formatMergedDatesHeader(groupedDates)
                    if (currentPageLineCount >= 45) {
                        addContinuationHeader(doc, station, formattedDatesHeader, subtitle)
                        currentPageLineCount = 2
                    }
                    addCenteredBoldHeader(doc, formattedDatesHeader)
                    currentPageLineCount++

                    val sampleDate = groupedDates.first()
                    for (rawLine in dateToCleanLines[sampleDate] ?: emptyList()) {
                        if (currentPageLineCount >= 45) {
                            addContinuationHeader(doc, station, formattedDatesHeader, subtitle)
                            currentPageLineCount = 2
                        }
                        val finalTrainLine = if (isMergedMultipleDays) nightNoteRegex.replace(rawLine, "").trim() else rawLine.trim()
                        val p = doc.createParagraph()
                        p.alignment = ParagraphAlignment.LEFT
                        p.spacingAfter = 0
                        p.spacingBefore = 0
                        p.spacingAfterLines = 0
                        p.spacingBeforeLines = 0
                        p.setSpacingBetween(1.0, LineSpacingRule.AUTO)
                        
                        val run = p.createRun()
                        run.fontFamily = "Times New Roman"
                        run.fontSize = 12
                        run.isBold = false
                        run.setText("        $finalTrainLine")
                        currentPageLineCount++
                    }
                }
                FileOutputStream(outFile).use { doc.write(it) }
                generatedCount++
            }
        }
        return Pair(generatedCount, targetFolder)
    }

    private fun addCenteredBoldHeader(doc: XWPFDocument, text: String) {
        val p = doc.createParagraph()
        p.alignment = ParagraphAlignment.CENTER
        p.spacingAfter = 0 
        p.spacingBefore = 0
        p.setSpacingBetween(1.0, LineSpacingRule.AUTO)
        val run = p.createRun()
        run.fontFamily = "Times New Roman"
        run.fontSize = 14
        run.isBold = true
        run.setText(text)
    }

    private fun addContinuationHeader(doc: XWPFDocument, station: String, datesHeader: String, subtitle: String) {
        val p1 = doc.createParagraph()
        p1.isPageBreak = true
        p1.alignment = ParagraphAlignment.CENTER
        p1.spacingAfter = 0
        p1.spacingBefore = 0
        p1.setSpacingBetween(1.0, LineSpacingRule.AUTO)
        val r1 = p1.createRun()
        r1.fontFamily = "Times New Roman"
        r1.fontSize = 14
        r1.isBold = true
        r1.setText("$station продолжение на $datesHeader")

        val p2 = doc.createParagraph()
        p2.alignment = ParagraphAlignment.CENTER
        p2.spacingAfter = 0
        p2.spacingBefore = 0
        p2.setSpacingBetween(1.0, LineSpacingRule.AUTO)
        val r2 = p2.createRun()
        r2.fontFamily = "Times New Roman"
        r2.fontSize = 14
        r2.isBold = true
        r2.setText(subtitle)
    }

    private fun formatMergedDatesHeader(dates: List<String>): String {
        if (dates.size == 1) return dates.first()
        val parsed = dates.map { d ->
            val parts = d.trim().split(Regex("\\s+"))
            Pair(parts.firstOrNull() ?: d, parts.drop(1).joinToString(" "))
        }
        return parsed.groupBy({ it.second }, { it.first }).entries.joinToString(", ") { (month, days) ->
            val sortedDays = days.sortedBy { it.toIntOrNull() ?: 99 }
            if (month.isNotBlank()) "${sortedDays.joinToString(",")} $month" else sortedDays.joinToString(",")
        }
    }
}

// ============================================================================
// 7. UI КОМПОНЕНТЫ
// ============================================================================

@OptIn(ExperimentalLayoutApi::class)
@Composable
fun TrainRApp() {
    val scope = rememberCoroutineScope()

    var currentNavTab: MainNavTab by remember { mutableStateOf(MainNavTab.UPLOADS) }
    var currentDirectionTab: DirectionTab by remember { mutableStateOf(DirectionTab.TO_MOSCOW) }
    var currentTrainFilter: TrainFilterType by remember { mutableStateOf(TrainFilterType.ALL) }
    var currentReportFilter: String by remember { mutableStateOf("ALL") }

    var searchQuery: String by remember { mutableStateOf("") }
    var systemStatusText: String by remember { mutableStateOf("Система готова. Полигон: Октябрьская ЖД / МЦД-3") }

    var afishaCacheMap: Map<String, AfishaBaseTrain> by remember { mutableStateOf(emptyMap()) }
    var stationScheduleMap: StationDateTrainMap by remember { mutableStateOf(emptyMap()) }
    var yandexTrainsList: List<YandexParsedTrain> by remember { mutableStateOf(emptyList()) }

    var loadedTelegramFiles by remember { mutableStateOf(listOf<File>()) }
    var loadedExportFiles by remember { mutableStateOf(listOf<File>()) }
    var loadedYandexCount by remember { mutableStateOf(0) }
    
    var showFileManagerDialog by remember { mutableStateOf(false) }

    var activePreviewStation: String by remember { mutableStateOf("Крюково") }
    var checkedStationsForWord: Set<String> by remember { mutableStateOf(setOf("Крюково", "Химки", "Сходня")) }
    var selectedFilterDates: Set<String> by remember { mutableStateOf(emptySet()) }

    var showCalendarDialog: Boolean by remember { mutableStateOf(false) }
    var selectedTrainForDossier: TrainChangeRow? by remember { mutableStateOf(null) }

    fun refreshData() {
        scope.launch(Dispatchers.IO) {
            var map = FileParsers.parseTelegramFiles(loadedTelegramFiles, afishaCacheMap)
            for (f in loadedExportFiles) {
                map = FileParsers.parseSummaryExportFile(f, afishaCacheMap, map)
            }
            stationScheduleMap = map
        }
    }

    LaunchedEffect(Unit) {
        if (FileParsers.cachedAfishaFile.exists()) {
            withContext(Dispatchers.IO) {
                val parsed = FileParsers.parseAfishaExcel(FileParsers.cachedAfishaFile)
                if (parsed.isNotEmpty()) {
                    afishaCacheMap = parsed
                    systemStatusText = "Загружен кэш базовой Афиши (${parsed.size} поездов)"
                }
            }
        }
    }

    val allDatesWithChanges: Set<String> = remember(stationScheduleMap) {
        stationScheduleMap.values.flatMap { it.keys }.toSet()
    }

    val baseRows: List<TrainChangeRow> = remember(
        stationScheduleMap,
        activePreviewStation,
        selectedFilterDates,
        currentDirectionTab,
        currentTrainFilter,
        searchQuery
    ) {
        val dateMap = stationScheduleMap[activePreviewStation] ?: emptyMap()
        val datesToInclude = if (selectedFilterDates.isEmpty()) dateMap.keys else dateMap.keys.intersect(selectedFilterDates)
        datesToInclude.flatMap { dt -> dateMap[dt]?.values ?: emptyList() }
            .filter { it.isToMoscow == currentDirectionTab.isToMoscow }
            .filter { row ->
                when (currentTrainFilter) {
                    TrainFilterType.ALL -> true
                    TrainFilterType.MCD -> row.isMcd
                    TrainFilterType.EXPRESS -> row.isExpress
                }
            }
            .filter { row ->
                if (searchQuery.isBlank()) true
                else row.trainNumber.contains(searchQuery.trim(), ignoreCase = true) ||
                        row.route.contains(searchQuery.trim(), ignoreCase = true) ||
                        row.actualChangesText.contains(searchQuery.trim(), ignoreCase = true)
            }
    }
    
    val reportRows = baseRows.filter {
        when(currentReportFilter) {
            "CANCELED" -> it.isCanceled
            "MCD" -> it.isMcd
            "MISMATCH" -> it.auditStatus != AuditStatus.OK
            else -> true
        }
    }

    Row(modifier = Modifier.fillMaxSize().background(TrainRTheme.BgLight)) {
        Column(
            modifier = Modifier
                .width(260.dp)
                .fillMaxHeight()
                .background(TrainRTheme.SidebarNavy)
                .padding(16.dp)
        ) {
            Column(modifier = Modifier.weight(1f).verticalScroll(rememberScrollState())) {
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    modifier = Modifier.padding(bottom = 24.dp)
                ) {
                    VintageTrainLogo(modifier = Modifier.size(44.dp))
                    Spacer(modifier = Modifier.width(12.dp))
                    Column {
                        Text("Train-R", color = Color.White, fontSize = 20.sp, fontWeight = FontWeight.ExtraBold)
                        Text("МТППК • МЦД-3", color = TrainRTheme.GoldAccent, fontSize = 11.sp, fontWeight = FontWeight.SemiBold)
                    }
                }

                HorizontalDivider(color = Color.White.copy(alpha = 0.12f))
                Spacer(modifier = Modifier.height(16.dp))

                SidebarNavButton(title = MainNavTab.UPLOADS.title, icon = Icons.Default.List, isSelected = currentNavTab == MainNavTab.UPLOADS) { currentNavTab = MainNavTab.UPLOADS }
                Spacer(modifier = Modifier.height(8.dp))
                SidebarNavButton(title = MainNavTab.REPORTS.title, icon = Icons.Default.Info, isSelected = currentNavTab == MainNavTab.REPORTS) { currentNavTab = MainNavTab.REPORTS }
                Spacer(modifier = Modifier.height(8.dp))
                SidebarNavButton(title = MainNavTab.YANDEX.title, icon = Icons.Default.CheckCircle, isSelected = currentNavTab == MainNavTab.YANDEX) { currentNavTab = MainNavTab.YANDEX }
            }

            Column(modifier = Modifier.wrapContentHeight(), verticalArrangement = Arrangement.spacedBy(10.dp)) {
                Button(
                    onClick = {
                        scope.launch(Dispatchers.IO) {
                            val (count, folder) = WordCancelGenerator.generateStationWordFiles(
                                selectedStations = checkedStationsForWord,
                                selectedDates = selectedFilterDates,
                                directionTab = currentDirectionTab,
                                allData = stationScheduleMap
                            )
                            systemStatusText = "Сформировано файлов Word (.docx): $count шт. Папка: ${folder.absolutePath}"
                        }
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue),
                    shape = RoundedCornerShape(8.dp),
                    modifier = Modifier.fillMaxWidth(),
                    contentPadding = PaddingValues(vertical = 12.dp)
                ) {
                    Icon(Icons.Default.Create, contentDescription = null, tint = Color.White, modifier = Modifier.size(16.dp))
                    Spacer(modifier = Modifier.width(6.dp))
                    Text("СОЗДАТЬ ОТМЕНЫ (${checkedStationsForWord.size} СТ.)", color = Color.White, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp)
                }

                OutlinedButton(
                    onClick = {
                        stationScheduleMap = emptyMap()
                        afishaCacheMap = emptyMap()
                        yandexTrainsList = emptyList()
                        loadedTelegramFiles = emptyList()
                        loadedExportFiles = emptyList()
                        loadedYandexCount = 0
                        selectedFilterDates = emptySet(); checkedStationsForWord = emptySet(); searchQuery = ""
                        scope.launch(Dispatchers.IO) { try { if (FileParsers.cachedAfishaFile.exists()) FileParsers.cachedAfishaFile.delete() } catch (e: Exception) {} }
                        systemStatusText = "Все данные сброшены, кэш очищен."
                    },
                    shape = RoundedCornerShape(8.dp),
                    colors = ButtonDefaults.outlinedButtonColors(contentColor = TrainRTheme.RedCancel),
                    border = BorderStroke(1.dp, TrainRTheme.RedCancel.copy(alpha = 0.5f)),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Icon(Icons.Default.Delete, contentDescription = null, modifier = Modifier.size(16.dp))
                    Spacer(modifier = Modifier.width(6.dp))
                    Text("СБРОСИТЬ И УДАЛИТЬ", fontWeight = FontWeight.Bold, fontSize = 11.sp)
                }
                
                Surface(
                    color = TrainRTheme.SidebarActive, shape = RoundedCornerShape(10.dp),
                    border = BorderStroke(1.dp, Color.White.copy(alpha = 0.1f)), modifier = Modifier.fillMaxWidth().padding(top = 8.dp)
                ) {
                    Column(modifier = Modifier.padding(12.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Box(modifier = Modifier.size(8.dp).clip(CircleShape).background(TrainRTheme.GreenSuccess))
                            Spacer(modifier = Modifier.width(8.dp))
                            Text("СТАТУС СИСТЕМЫ", color = Color.White.copy(alpha = 0.7f), fontSize = 10.sp, fontWeight = FontWeight.Bold)
                        }
                        Spacer(modifier = Modifier.height(6.dp))
                        Text(systemStatusText, color = Color.White, fontSize = 11.sp, lineHeight = 14.sp)
                    }
                }
            }
        }

        Column(modifier = Modifier.weight(1f).fillMaxHeight()) {
            Surface(color = Color.White, shadowElevation = 2.dp, modifier = Modifier.fillMaxWidth()) {
                Row(
                    modifier = Modifier.fillMaxWidth().padding(horizontal = 20.dp, vertical = 12.dp),
                    verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.SpaceBetween
                ) {
                    OutlinedTextField(
                        value = searchQuery, onValueChange = { searchQuery = it },
                        placeholder = { Text("Поиск поезда...", fontSize = 13.sp, color = TrainRTheme.TextSub) },
                        leadingIcon = { Icon(Icons.Default.Search, contentDescription = null, tint = TrainRTheme.TextSub) },
                        trailingIcon = { if (searchQuery.isNotEmpty()) IconButton(onClick = { searchQuery = "" }) { Icon(Icons.Default.Clear, contentDescription = null) } },
                        singleLine = true, shape = RoundedCornerShape(8.dp), modifier = Modifier.width(280.dp).height(50.dp)
                    )

                    Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                        OutlinedButton(
                            onClick = { showFileManagerDialog = true },
                            shape = RoundedCornerShape(8.dp),
                            border = BorderStroke(1.dp, TrainRTheme.NavyDark),
                            contentPadding = PaddingValues(horizontal = 12.dp, vertical = 8.dp)
                        ) {
                            Icon(Icons.Default.Settings, contentDescription = null, tint = TrainRTheme.NavyDark, modifier = Modifier.size(16.dp))
                            Spacer(modifier = Modifier.width(6.dp))
                            Text("ФАЙЛЫ (${loadedTelegramFiles.size + loadedExportFiles.size})", color = TrainRTheme.NavyDark, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }

                        TopBarLoadButton("АФИША", afishaCacheMap.size, TrainRTheme.RoyalBlue) {
                            val f = pickFilesDialog("Афиша .xlsx", arrayOf("xlsx", "xls"), false).firstOrNull()
                            if (f != null) scope.launch(Dispatchers.IO) { afishaCacheMap = FileParsers.parseAfishaExcel(FileParsers.ensureCacheAfisha(f)); systemStatusText = "Афиша закэширована" }
                        }
                        TopBarLoadButton("ТЕЛЕГРАММА", loadedTelegramFiles.size, TrainRTheme.AmberWarn) {
                            val files = pickFilesDialog("Телеграммы", arrayOf("docx", "pdf", "txt"), true)
                            if (files.isNotEmpty()) {
                                loadedTelegramFiles = loadedTelegramFiles + files
                                systemStatusText = "Добавлены телеграммы. Пересчет..."
                                refreshData()
                            }
                        }
                        TopBarLoadButton("ВЫГРУЗКА", loadedExportFiles.size, TrainRTheme.GreenSuccess) {
                            val files = pickFilesDialog("Выгрузки", arrayOf("txt", "docx"), true)
                            if (files.isNotEmpty()) {
                                loadedExportFiles = loadedExportFiles + files
                                systemStatusText = "Добавлены выгрузки. Пересчет..."
                                refreshData()
                            }
                        }
                    }
                }
            }

            Surface(
                color = Color.White, border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                modifier = Modifier.fillMaxWidth().padding(horizontal = 20.dp, vertical = 10.dp), shape = RoundedCornerShape(10.dp)
            ) {
                Column(modifier = Modifier.padding(14.dp)) {
                    Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                        Row(horizontalArrangement = Arrangement.spacedBy(10.dp), verticalAlignment = Alignment.CenterVertically) {
                            OutlinedButton(onClick = { showCalendarDialog = true }, shape = RoundedCornerShape(8.dp)) {
                                Icon(Icons.Default.List, contentDescription = null, modifier = Modifier.size(18.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text(if (selectedFilterDates.isEmpty()) "ФИЛЬТР ПО ДАТЕ (ВСЕ)" else "ДАТЫ: Выбрано (${selectedFilterDates.size})", fontWeight = FontWeight.Bold, fontSize = 12.sp)
                            }
                            OutlinedButton(
                                onClick = { systemStatusText = "Выполнена сверка."; currentNavTab = MainNavTab.REPORTS },
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(18.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("СВЕРИТЬ ВСЕ", fontWeight = FontWeight.Bold, fontSize = 12.sp)
                            }
                        }
                        if (selectedFilterDates.isNotEmpty()) {
                            TextButton(onClick = { selectedFilterDates = emptySet() }) {
                                Text("Сбросить фильтр дат", fontSize = 12.sp, color = TrainRTheme.RedCancel)
                            }
                        }
                    }
                    Spacer(modifier = Modifier.height(10.dp))
                    HorizontalDivider(color = TrainRTheme.BorderCol)
                    Spacer(modifier = Modifier.height(8.dp))

                    Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                        Text("Станции полигона (чекбокс — генерация Word, клик по названию — предпросмотр):", fontSize = 11.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.SemiBold)
                        Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                            TextButton(onClick = { checkedStationsForWord = PolygonConfig.ALL_STATIONS.toSet() }, contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)) { Text("Выбрать все", fontSize = 11.sp) }
                            TextButton(onClick = { checkedStationsForWord = PolygonConfig.MCD3_STATIONS.intersect(PolygonConfig.ALL_STATIONS.toSet()) }, contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)) { Text("Только МЦД-3", fontSize = 11.sp) }
                            TextButton(onClick = { checkedStationsForWord = emptySet() }, contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)) { Text("Снять все", fontSize = 11.sp) }
                        }
                    }
                    FlowRow(modifier = Modifier.fillMaxWidth().padding(top = 4.dp), horizontalArrangement = Arrangement.spacedBy(6.dp), verticalArrangement = Arrangement.spacedBy(6.dp)) {
                        for (st in PolygonConfig.ALL_STATIONS) {
                            val isChecked = st in checkedStationsForWord
                            val isPreview = st == activePreviewStation
                            val hasChanges = (stationScheduleMap[st]?.isNotEmpty() == true)

                            Surface(
                                shape = RoundedCornerShape(6.dp),
                                color = when { isPreview -> TrainRTheme.ActionBlue.copy(alpha = 0.14f)
                                    hasChanges -> Color(0xFFF1F5F9) else -> Color.White },
                                border = BorderStroke(if (isPreview) 1.5.dp else 1.dp, if (isPreview) TrainRTheme.ActionBlue else TrainRTheme.BorderCol)
                            ) {
                                Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.clickable { activePreviewStation = st }.padding(end = 8.dp, top = 2.dp, bottom = 2.dp)) {
                                    Checkbox(checked = isChecked, onCheckedChange = { checked -> checkedStationsForWord = if (checked) checkedStationsForWord + st else checkedStationsForWord - st }, modifier = Modifier.size(26.dp))
                                    Text(st, fontSize = 11.sp, fontWeight = if (isPreview || hasChanges) FontWeight.Bold else FontWeight.Normal, color = if (isPreview) TrainRTheme.RoyalBlue else TrainRTheme.TextDark)
                                }
                            }
                        }
                    }
                }
            }

            Box(modifier = Modifier.weight(1f).fillMaxWidth().padding(horizontal = 20.dp, vertical = 6.dp)) {
                when (currentNavTab) {
                    MainNavTab.UPLOADS -> TrainTableSection(activeStation = activePreviewStation, directionTab = currentDirectionTab, onDirectionChange = { currentDirectionTab = it }, trainFilter = currentTrainFilter, onTrainFilterChange = { currentTrainFilter = it }, rows = baseRows, onRowClick = { selectedTrainForDossier = it })
                    MainNavTab.REPORTS -> StationAnalyticsReportView(activeStation = activePreviewStation, baseRows = baseRows, reportRows = reportRows, currentFilter = currentReportFilter, onFilterChange = { currentReportFilter = it }, onRowClick = { selectedTrainForDossier = it })
                    MainNavTab.YANDEX -> YandexAuditTabView(activeStation = activePreviewStation, rows = baseRows, yandexList = yandexTrainsList, onRowClick = { selectedTrainForDossier = it })
                }
            }
        }
    }

    if (showFileManagerDialog) {
        FileManagerDialog(
            telegrams = loadedTelegramFiles,
            exports = loadedExportFiles,
            onRemoveTelegram = { file -> 
                loadedTelegramFiles = loadedTelegramFiles - file 
                refreshData()
            },
            onRemoveExport = { file -> 
                loadedExportFiles = loadedExportFiles - file 
                refreshData()
            },
            onDismiss = { showFileManagerDialog = false }
        )
    }

    if (showCalendarDialog) CalendarGridPickerDialog(allDatesWithChanges, selectedFilterDates, { dt -> selectedFilterDates = if (dt in selectedFilterDates) selectedFilterDates - dt else selectedFilterDates + dt }, { selectedFilterDates = emptySet() }, { showCalendarDialog = false })
    selectedTrainForDossier?.let { TrainDossierDialog(it) { selectedTrainForDossier = null } }
}

@Composable
fun FileManagerDialog(
    telegrams: List<File>,
    exports: List<File>,
    onRemoveTelegram: (File) -> Unit,
    onRemoveExport: (File) -> Unit,
    onDismiss: () -> Unit
) {
    Dialog(onDismissRequest = onDismiss, properties = DialogProperties(usePlatformDefaultWidth = false)) {
        Surface(color = Color.White, shape = RoundedCornerShape(12.dp), modifier = Modifier.width(600.dp).heightIn(max = 600.dp)) {
            Column(modifier = Modifier.fillMaxWidth().padding(20.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Text("МЕНЕДЖЕР ЗАГРУЖЕННЫХ ФАЙЛОВ", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, contentDescription = "Закрыть") }
                }
                Spacer(modifier = Modifier.height(12.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                
                Column(modifier = Modifier.weight(1f, fill = false).verticalScroll(rememberScrollState()).padding(vertical = 12.dp)) {
                    Text("Телеграммы (${telegrams.size})", fontWeight = FontWeight.Bold, color = TrainRTheme.AmberWarn, fontSize = 13.sp)
                    if (telegrams.isEmpty()) Text("Нет загруженных телеграмм", fontSize = 12.sp, color = TrainRTheme.TextSub)
                    telegrams.forEach { file ->
                        Row(modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                            Text(file.name, fontSize = 12.sp, maxLines = 1, overflow = TextOverflow.Ellipsis, modifier = Modifier.weight(1f))
                            IconButton(onClick = { onRemoveTelegram(file) }, modifier = Modifier.size(24.dp)) {
                                Icon(Icons.Default.Delete, contentDescription = "Удалить", tint = TrainRTheme.RedCancel, modifier = Modifier.size(18.dp))
                            }
                        }
                    }
                    
                    Spacer(modifier = Modifier.height(16.dp))
                    
                    Text("Выгрузки (${exports.size})", fontWeight = FontWeight.Bold, color = TrainRTheme.GreenSuccess, fontSize = 13.sp)
                    if (exports.isEmpty()) Text("Нет загруженных выгрузок", fontSize = 12.sp, color = TrainRTheme.TextSub)
                    exports.forEach { file ->
                        Row(modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                            Text(file.name, fontSize = 12.sp, maxLines = 1, overflow = TextOverflow.Ellipsis, modifier = Modifier.weight(1f))
                            IconButton(onClick = { onRemoveExport(file) }, modifier = Modifier.size(24.dp)) {
                                Icon(Icons.Default.Delete, contentDescription = "Удалить", tint = TrainRTheme.RedCancel, modifier = Modifier.size(18.dp))
                            }
                        }
                    }
                }
                
                Spacer(modifier = Modifier.height(12.dp))
                Button(onClick = onDismiss, modifier = Modifier.fillMaxWidth(), colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark)) {
                    Text("ЗАКРЫТЬ", color = Color.White, fontWeight = FontWeight.Bold)
                }
            }
        }
    }
}

@Composable
fun VintageTrainLogo(modifier: Modifier = Modifier) {
    Canvas(modifier = modifier) {
        val w = size.width
        val h = size.height
        drawRoundRect(color = TrainRTheme.ActionBlue, topLeft = Offset(0f, 0f), size = Size(w, h), cornerRadius = CornerRadius(12f, 12f))
        val bodyPath = Path().apply { moveTo(w * 0.22f, h * 0.20f); lineTo(w * 0.78f, h * 0.20f); lineTo(w * 0.85f, h * 0.72f); lineTo(w * 0.15f, h * 0.72f); close() }
        drawPath(path = bodyPath, color = Color.White)
        drawRoundRect(color = TrainRTheme.NavyDark, topLeft = Offset(w * 0.28f, h * 0.27f), size = Size(w * 0.44f, h * 0.20f), cornerRadius = CornerRadius(4f, 4f))
        drawCircle(color = TrainRTheme.GoldAccent, radius = w * 0.05f, center = Offset(w * 0.50f, h * 0.58f))
        drawLine(color = TrainRTheme.GoldAccent, start = Offset(w * 0.15f, h * 0.82f), end = Offset(w * 0.85f, h * 0.82f), strokeWidth = 3f)
    }
}

@Composable
fun SidebarNavButton(title: String, icon: androidx.compose.ui.graphics.vector.ImageVector, isSelected: Boolean, onClick: () -> Unit) {
    Surface(color = if (isSelected) TrainRTheme.ActionBlue else Color.Transparent, shape = RoundedCornerShape(8.dp), modifier = Modifier.fillMaxWidth().clickable { onClick() }) {
        Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.padding(horizontal = 12.dp, vertical = 12.dp)) {
            Icon(imageVector = icon, contentDescription = title, tint = if (isSelected) Color.White else Color.White.copy(alpha = 0.7f), modifier = Modifier.size(20.dp))
            Spacer(modifier = Modifier.width(10.dp))
            Text(text = title, color = if (isSelected) Color.White else Color.White.copy(alpha = 0.75f), fontSize = 12.sp, fontWeight = if (isSelected) FontWeight.ExtraBold else FontWeight.Medium)
        }
    }
}

@Composable
fun TopBarLoadButton(label: String, count: Int, badgeColor: Color, onClick: () -> Unit) {
    OutlinedButton(onClick = onClick, shape = RoundedCornerShape(8.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), contentPadding = PaddingValues(horizontal = 12.dp, vertical = 8.dp)) {
        Icon(imageVector = Icons.Default.Add, contentDescription = label, tint = TrainRTheme.TextDark, modifier = Modifier.size(16.dp))
        Spacer(modifier = Modifier.width(6.dp))
        Text(text = label, color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold, fontSize = 12.sp)
        Spacer(modifier = Modifier.width(6.dp))
        Surface(color = badgeColor, shape = RoundedCornerShape(10.dp)) { Text(text = count.toString(), color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp)) }
    }
}

@Composable
fun FilterChipItem(label: String, isSelected: Boolean, onClick: () -> Unit) {
    Surface(color = if (isSelected) TrainRTheme.NavyDark else Color.White, shape = RoundedCornerShape(16.dp), border = BorderStroke(1.dp, if (isSelected) TrainRTheme.NavyDark else TrainRTheme.BorderCol), modifier = Modifier.clickable { onClick() }) {
        Text(text = label, color = if (isSelected) Color.White else TrainRTheme.TextDark, fontSize = 12.sp, fontWeight = FontWeight.SemiBold, modifier = Modifier.padding(horizontal = 14.dp, vertical = 6.dp))
    }
}

@Composable
fun TrainTableSection(activeStation: String, directionTab: DirectionTab, onDirectionChange: (DirectionTab) -> Unit, trainFilter: TrainFilterType, onTrainFilterChange: (TrainFilterType) -> Unit, rows: List<TrainChangeRow>, onRowClick: (TrainChangeRow) -> Unit) {
    Surface(color = Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = Modifier.fillMaxSize()) {
        Column(modifier = Modifier.fillMaxSize()) {
            Row(modifier = Modifier.fillMaxWidth().padding(14.dp), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    DirectionTab.entries.forEach { tab ->
                        val selected = tab == directionTab
                        Button(onClick = { onDirectionChange(tab) }, colors = ButtonDefaults.buttonColors(containerColor = if (selected) TrainRTheme.RoyalBlue else Color(0xFFF1F5F9), contentColor = if (selected) Color.White else TrainRTheme.TextDark), shape = RoundedCornerShape(8.dp)) { Text(tab.title, fontWeight = FontWeight.Bold, fontSize = 12.sp) }
                    }
                }
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp), verticalAlignment = Alignment.CenterVertically) {
                    Text("Станция: $activeStation • Составы:", fontSize = 12.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.Medium)
                    TrainFilterType.entries.forEach { filter -> FilterChipItem(filter.label, trainFilter == filter) { onTrainFilterChange(filter) } }
                }
            }
            if (activeStation == "Левобережная" && !directionTab.isToMoscow) {
                Box(modifier = Modifier.fillMaxWidth().background(TrainRTheme.AmberWarn.copy(alpha = 0.15f)).padding(12.dp)) { Text("На ст. Левобережная платформа №2 (путь №2 «от Москвы») закрыта. Изменения не выводятся.", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold, fontSize = 12.sp) }
            }
            Row(modifier = Modifier.fillMaxWidth().background(TrainRTheme.NavyDark).padding(horizontal = 12.dp, vertical = 10.dp), verticalAlignment = Alignment.CenterVertically) {
                Text("ДАТЫ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.9f))
                Text("№ ПОЕЗДА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
                Text("МАРШРУТ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.3f))
                Text("АФИША", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.7f))
                Text("ТЕЛЕГР.", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.8f))
                Text("ВЫГРУЗКА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.8f))
                Text("ИЗМЕНЕНИЯ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(2.3f))
                Text("СВЕРКА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
            }
            if (rows.isEmpty()) { Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) { Text("Нет данных об изменениях по станции «$activeStation».", color = TrainRTheme.TextSub, fontSize = 14.sp) }
            } else {
                LazyColumn(modifier = Modifier.fillMaxSize()) {
                    items(rows.sortedBy { getSortingTime(it) }, key = { it.id }) { row ->
                        Row(modifier = Modifier.fillMaxWidth().background(if (row.globalErrorNote != null) TrainRTheme.RedCancel.copy(alpha = 0.14f) else if (row.isCanceled) TrainRTheme.RedCancel.copy(alpha = 0.06f) else Color.White).clickable { onRowClick(row) }.padding(horizontal = 12.dp, vertical = 10.dp), verticalAlignment = Alignment.CenterVertically) {
                            Text(row.dateStr, fontSize = 12.sp, fontWeight = FontWeight.SemiBold, color = TrainRTheme.TextDark, modifier = Modifier.weight(0.9f))
                            Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.weight(1.1f)) {
                                Text("№${row.trainNumber}", fontSize = 13.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.RoyalBlue)
                                if (row.isMcd) { Spacer(Modifier.width(6.dp)); Surface(color = TrainRTheme.PurpleMcd, shape = RoundedCornerShape(4.dp)) { Text("МЦД", color = Color.White, fontSize = 9.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(horizontal = 4.dp, vertical = 1.dp)) } }
                            }
                            Text(row.route, fontSize = 12.sp, color = TrainRTheme.TextDark, maxLines = 1, overflow = TextOverflow.Ellipsis, modifier = Modifier.weight(1.3f))
                            Text(row.afishaTime, fontSize = 12.sp, fontFamily = FontFamily.Monospace, color = TrainRTheme.TextSub, modifier = Modifier.weight(0.7f))
                            
                            val isTgCanceled = row.isCanceled || row.telegramTime.contains("ОТМЕН", true)
                            val isExpCanceled = row.isCanceled || row.exportTime.contains("ОТМЕН", true)
                            
                            Text(row.telegramTime, fontSize = 12.sp, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Bold, color = if (isTgCanceled) TrainRTheme.RedCancel else TrainRTheme.GreenSuccess, modifier = Modifier.weight(0.8f))
                            Text(row.exportTime, fontSize = 12.sp, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Bold, color = if (isExpCanceled) TrainRTheme.RedCancel else TrainRTheme.TextSub, modifier = Modifier.weight(0.8f))
                            Column(modifier = Modifier.weight(2.3f).padding(end = 8.dp)) {
                                Text(row.actualChangesText, fontSize = 12.sp, color = TrainRTheme.TextDark)
                                if (row.globalErrorNote != null) Text(row.globalErrorNote, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RedCancel)
                            }
                            Surface(color = row.auditStatus.color.copy(alpha = 0.14f), shape = RoundedCornerShape(6.dp), border = BorderStroke(1.dp, row.auditStatus.color), modifier = Modifier.weight(1.1f)) {
                                Text(row.auditStatus.label, color = row.auditStatus.color, fontSize = 10.sp, fontWeight = FontWeight.ExtraBold, textAlign = TextAlign.Center, modifier = Modifier.padding(horizontal = 6.dp, vertical = 4.dp))
                            }
                        }
                        HorizontalDivider(color = TrainRTheme.BorderCol)
                    }
                }
            }
        }
    }
}

@Composable
fun StationAnalyticsReportView(activeStation: String, baseRows: List<TrainChangeRow>, reportRows: List<TrainChangeRow>, currentFilter: String, onFilterChange: (String) -> Unit, onRowClick: (TrainChangeRow) -> Unit) {
    val totalCount = baseRows.size; val canceledCount = baseRows.count { it.isCanceled }; val mcdCount = baseRows.count { it.isMcd }; val mismatchCount = baseRows.count { it.auditStatus != AuditStatus.OK }
    Column(modifier = Modifier.fillMaxSize()) {
        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(12.dp)) {
            StatSummaryBox("ВСЕГО ИЗМЕНЕНИЙ ($activeStation)", totalCount.toString(), TrainRTheme.RoyalBlue, currentFilter == "ALL", { onFilterChange("ALL") }, Modifier.weight(1f))
            StatSummaryBox("ПОЛНЫЕ ОТМЕНЫ ПОЕЗДОВ", canceledCount.toString(), TrainRTheme.RedCancel, currentFilter == "CANCELED", { onFilterChange("CANCELED") }, Modifier.weight(1f))
            StatSummaryBox("ПОЕЗДА МЦД-3 (ИВОЛГА/ЭП2Д)", mcdCount.toString(), TrainRTheme.PurpleMcd, currentFilter == "MCD", { onFilterChange("MCD") }, Modifier.weight(1f))
            StatSummaryBox("РАСХОЖДЕНИЯ 4 ИСТОЧНИКОВ", mismatchCount.toString(), TrainRTheme.AmberWarn, currentFilter == "MISMATCH", { onFilterChange("MISMATCH") }, Modifier.weight(1f))
        }
        Spacer(modifier = Modifier.height(12.dp))
        Surface(color = Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = Modifier.fillMaxSize()) {
            Column(modifier = Modifier.fillMaxSize().padding(14.dp)) {
                Text("СКВОЗНАЯ 4-СТОРОННЯЯ СВЕРКА • СТ. $activeStation", fontSize = 14.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                Spacer(modifier = Modifier.height(10.dp))
                LazyColumn(modifier = Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(8.dp)) { items(reportRows, key = { it.id }) { row -> DetailedReportTrainCard(row, { onRowClick(row) }) } }
            }
        }
    }
}

@Composable
fun StatSummaryBox(title: String, value: String, accentColor: Color, isSelected: Boolean, onClick: () -> Unit, modifier: Modifier = Modifier) {
    Surface(color = if (isSelected) accentColor.copy(alpha = 0.1f) else Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(if (isSelected) 2.dp else 1.dp, if (isSelected) accentColor else TrainRTheme.BorderCol), modifier = modifier.clickable { onClick() }) {
        Column(modifier = Modifier.padding(16.dp)) { Text(title, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub); Spacer(modifier = Modifier.height(6.dp)); Text(value, fontSize = 26.sp, fontWeight = FontWeight.ExtraBold, color = accentColor) }
    }
}

@Composable
fun DetailedReportTrainCard(row: TrainChangeRow, onClick: () -> Unit) {
    Surface(color = Color(0xFFF8FAFC), shape = RoundedCornerShape(8.dp), border = BorderStroke(1.dp, row.auditStatus.color.copy(alpha = 0.5f)), modifier = Modifier.fillMaxWidth().clickable { onClick() }) {
        Column(modifier = Modifier.padding(12.dp)) {
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text("Поезд №${row.trainNumber} (${row.route})", fontWeight = FontWeight.ExtraBold, fontSize = 14.sp, color = TrainRTheme.NavyDark)
                    Spacer(modifier = Modifier.width(10.dp))
                    Text("Дата: ${row.dateStr} • Основание: ${row.telegramBasis}", fontSize = 12.sp, color = TrainRTheme.TextSub)
                }
                Surface(color = row.auditStatus.color, shape = RoundedCornerShape(6.dp)) { Text(row.auditStatus.label, color = Color.White, fontSize = 10.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp)) }
            }
            Spacer(modifier = Modifier.height(6.dp))
            Text("Цепочка сверки: Афиша ${row.afishaTime} ➔ Телеграмма ${row.telegramTime} ➔ Выгрузка ${row.exportTime} ➔ Яндекс ${row.yandexTime}", fontSize = 12.sp, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.SemiBold, color = TrainRTheme.RoyalBlue)
            Spacer(modifier = Modifier.height(4.dp))
            Text("Строка для Word: ${row.wordExportLine}", fontSize = 12.sp, color = TrainRTheme.TextDark)
            Spacer(modifier = Modifier.height(4.dp))
            Text("Текст источника: ${row.rawTelegramSnippet}", fontSize = 11.sp, color = TrainRTheme.TextSub, fontFamily = FontFamily.Monospace)
        }
    }
}

@Composable
fun YandexAuditTabView(activeStation: String, rows: List<TrainChangeRow>, yandexList: List<YandexParsedTrain>, onRowClick: (TrainChangeRow) -> Unit) {
    Surface(color = Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = Modifier.fillMaxSize()) {
        Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                Column { Text("АУДИТ PDF-РАСПИСАНИЯ ЯНДЕКСА ПРОТИВ ТЕЛЕГРАММ РЖД И БАЗОВОЙ АФИШИ", fontSize = 15.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark); Text("Станция контроля: $activeStation • Загружено записей из PDF Яндекса: ${yandexList.size}", fontSize = 12.sp, color = TrainRTheme.TextSub) }
                Surface(color = TrainRTheme.YandexRed.copy(alpha = 0.12f), shape = RoundedCornerShape(8.dp), border = BorderStroke(1.dp, TrainRTheme.YandexRed)) { Text("Контроль Яндекс.Расписаний", color = TrainRTheme.YandexRed, fontWeight = FontWeight.Bold, fontSize = 12.sp, modifier = Modifier.padding(horizontal = 12.dp, vertical = 6.dp)) }
            }
            Spacer(modifier = Modifier.height(12.dp))
            HorizontalDivider(color = TrainRTheme.BorderCol)
            Spacer(modifier = Modifier.height(12.dp))
            LazyColumn(modifier = Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(8.dp)) { items(rows, key = { it.id }) { row -> DetailedReportTrainCard(row, { onRowClick(row) }) } }
        }
    }
}

@Composable
fun TrainDossierDialog(trainRow: TrainChangeRow, onDismiss: () -> Unit) {
    Dialog(onDismissRequest = onDismiss, properties = DialogProperties(usePlatformDefaultWidth = false)) {
        Surface(color = Color.White, shape = RoundedCornerShape(12.dp), modifier = Modifier.width(960.dp).heightIn(min = 520.dp, max = 800.dp)) {
            Column(modifier = Modifier.fillMaxSize().padding(20.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column { Text("ПОЛНЫЙ ОТЧЕТ ПО ПОЕЗДУ (ДОСЬЕ) — №${trainRow.trainNumber}", fontSize = 18.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark); Text("${trainRow.route} • Дата: ${trainRow.dateStr} • ${trainRow.telegramBasis}", fontSize = 12.sp, color = TrainRTheme.TextSub) }
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, contentDescription = "Закрыть") }
                }
                Spacer(modifier = Modifier.height(12.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                Spacer(modifier = Modifier.height(12.dp))
                Text("Нитка следования по станциям (Афиша ➔ Телеграмма ➔ Выгрузка ➔ Разница):", fontWeight = FontWeight.Bold, fontSize = 13.sp, color = TrainRTheme.NavyDark)
                Spacer(modifier = Modifier.height(8.dp))

                val stopsToShow = if (trainRow.fullThread.isNotEmpty()) trainRow.fullThread else listOf(StationStopEntry(trainRow.station, trainRow.afishaTime, trainRow.telegramTime, trainRow.exportTime, trainRow.yandexTime))
                Column(modifier = Modifier.fillMaxWidth().border(1.dp, TrainRTheme.BorderCol, RoundedCornerShape(8.dp))) {
                    Row(modifier = Modifier.fillMaxWidth().background(TrainRTheme.NavyDark).padding(10.dp)) {
                        Text("СТАНЦИЯ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.5f))
                        Text("АФИША", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(0.9f))
                        Text("ТЕЛЕГРАММА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.1f))
                        Text("ВЫГРУЗКА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.1f))
                        Text("РАЗНИЦА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(0.9f))
                        Text("СТАТУС ОСТАНОВКИ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.5f))
                    }
                    for (stop in stopsToShow) {
                        val activeComparisonTime = stop.exportTime.takeIf { it != "—" } ?: stop.telegramTime
                        val diffStr = calculateMinutesDiff(stop.afishaTime, activeComparisonTime)
                        
                        val isTgCanceled = stop.isSkippedByTelegram || stop.telegramTime.contains("ОТМЕН", true)
                        val isExpCanceled = stop.exportTime.contains("ОТМЕН", true)
                        
                        Row(modifier = Modifier.fillMaxWidth().padding(horizontal = 10.dp, vertical = 8.dp), verticalAlignment = Alignment.CenterVertically) {
                            Text(stop.stationName, fontWeight = FontWeight.Bold, fontSize = 12.sp, modifier = Modifier.weight(1.5f))
                            Text(stop.afishaTime, fontFamily = FontFamily.Monospace, fontSize = 12.sp, modifier = Modifier.weight(0.9f))
                            Text(stop.telegramTime, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Bold, color = if (isTgCanceled) TrainRTheme.RedCancel else TrainRTheme.TextDark, fontSize = 12.sp, modifier = Modifier.weight(1.1f))
                            Text(stop.exportTime, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Bold, color = if (isExpCanceled) TrainRTheme.RedCancel else TrainRTheme.TextSub, fontSize = 12.sp, modifier = Modifier.weight(1.1f))
                            Text(diffStr, fontFamily = FontFamily.Monospace, fontSize = 12.sp, modifier = Modifier.weight(0.9f))
                            Text(when { isExpCanceled || isTgCanceled -> "БЕЗ ОСТАНОВКИ (прослед.)"; stop.isAddedByTelegram -> "НАЗНАЧЕНА ОСТАНОВКА"; stop.isTechStop -> "ТЕХ. СТОЯНКА (без пасс.)"; else -> "Посадка / высадка пасс." }, fontSize = 11.sp, color = if (isExpCanceled || isTgCanceled) TrainRTheme.RedCancel else TrainRTheme.TextDark, modifier = Modifier.weight(1.5f))
                        }
                    }
                }
                
                Spacer(modifier = Modifier.height(16.dp))
                
                Row(modifier = Modifier.weight(1f).fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(16.dp)) {
                    Column(modifier = Modifier.weight(1f)) {
                        Text("Выдержка из Телеграммы:", fontWeight = FontWeight.Bold, fontSize = 13.sp, color = TrainRTheme.NavyDark)
                        Spacer(modifier = Modifier.height(6.dp))
                        Surface(color = Color(0xFFF1F5F9), shape = RoundedCornerShape(8.dp), modifier = Modifier.fillMaxSize()) { 
                            Text(trainRow.rawTelegramSnippet.ifBlank { "Телеграмма для данного поезда на эту дату не загружена." }, fontSize = 12.sp, fontFamily = FontFamily.Monospace, modifier = Modifier.padding(12.dp).verticalScroll(rememberScrollState())) 
                        }
                    }
                    Column(modifier = Modifier.weight(1f)) {
                        Text("Выдержка из Выгрузки:", fontWeight = FontWeight.Bold, fontSize = 13.sp, color = TrainRTheme.NavyDark)
                        Spacer(modifier = Modifier.height(6.dp))
                        Surface(color = Color(0xFFF1F5F9), shape = RoundedCornerShape(8.dp), modifier = Modifier.fillMaxSize()) { 
                            Text(trainRow.rawExportSnippet.ifBlank { "Выгрузка для данного поезда на эту дату не загружена." }, fontSize = 12.sp, fontFamily = FontFamily.Monospace, modifier = Modifier.padding(12.dp).verticalScroll(rememberScrollState())) 
                        }
                    }
                }
                
                Spacer(modifier = Modifier.height(16.dp))
                
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween) {
                    Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.BorderCol)) { Text("Свернуть", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold) }
                    Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark)) { Text("Закрыть досье", color = Color.White, fontWeight = FontWeight.Bold) }
                }
            }
        }
    }
}

@Composable
fun CalendarGridPickerDialog(availableDatesWithChanges: Set<String>, selectedDates: Set<String>, onToggleDate: (String) -> Unit, onClearAll: () -> Unit, onDismiss: () -> Unit) {
    val currentMonth = remember { YearMonth.of(2026, 10) }
    val nextMonth = remember { currentMonth.plusMonths(1) }
    Dialog(onDismissRequest = onDismiss, properties = DialogProperties(usePlatformDefaultWidth = false)) {
        Surface(color = Color.White, shape = RoundedCornerShape(12.dp), modifier = Modifier.width(720.dp).padding(16.dp)) {
            Column(modifier = Modifier.padding(20.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column { Text("КАЛЕНДАРЬ ОПЕРАТИВНЫХ ИЗМЕНЕНИЙ (НА 2 МЕСЯЦА ВПЕРЁД)", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark); Text("Кликабельны только дни, в которых обнаружены изменения в расписании", fontSize = 12.sp, color = TrainRTheme.TextSub) }
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, contentDescription = "Закрыть") }
                }
                Spacer(modifier = Modifier.height(16.dp))
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(20.dp)) {
                    MonthGridCalendar(currentMonth, "октября", "Октябрь 2026", availableDatesWithChanges, selectedDates, onToggleDate, Modifier.weight(1f))
                    MonthGridCalendar(nextMonth, "ноября", "Ноябрь 2026", availableDatesWithChanges, selectedDates, onToggleDate, Modifier.weight(1f))
                }
                Spacer(modifier = Modifier.height(18.dp))
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    TextButton(onClick = onClearAll) { Text("Сбросить выделение (показать все дни)", color = TrainRTheme.RedCancel) }
                    Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.BorderCol)) { Text("Свернуть", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold) }
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue)) { Text("Применить (${selectedDates.size})", color = Color.White, fontWeight = FontWeight.Bold) }
                    }
                }
            }
        }
    }
}

@Composable
fun MonthGridCalendar(yearMonth: YearMonth, monthGenitiveName: String, monthTitle: String, availableDatesWithChanges: Set<String>, selectedDates: Set<String>, onToggleDate: (String) -> Unit, modifier: Modifier = Modifier) {
    Surface(color = Color(0xFFF8FAFC), shape = RoundedCornerShape(8.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = modifier) {
        Column(modifier = Modifier.padding(12.dp)) {
            Text(monthTitle, fontWeight = FontWeight.ExtraBold, fontSize = 14.sp, color = TrainRTheme.NavyDark, modifier = Modifier.padding(bottom = 8.dp))
            val daysInMonth = yearMonth.lengthOfMonth()
            val rows = (1..daysInMonth).chunked(7)
            for (week in rows) {
                Row(modifier = Modifier.fillMaxWidth().padding(vertical = 3.dp), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                    for (day in week) {
                        val dateKey = "$day $monthGenitiveName"; val isClickable = dateKey in availableDatesWithChanges; val isSelected = dateKey in selectedDates
                        Box(contentAlignment = Alignment.Center, modifier = Modifier.size(36.dp).clip(RoundedCornerShape(6.dp)).background(when { isSelected -> TrainRTheme.ActionBlue; isClickable -> TrainRTheme.GoldAccent.copy(alpha = 0.25f); else -> Color.Transparent }).border(1.dp, when { isSelected -> TrainRTheme.ActionBlue; isClickable -> TrainRTheme.GoldAccent; else -> TrainRTheme.BorderCol }, RoundedCornerShape(6.dp)).let { mod -> if (isClickable) mod.clickable { onToggleDate(dateKey) } else mod }) {
                            Text(day.toString(), fontSize = 12.sp, fontWeight = if (isClickable || isSelected) FontWeight.ExtraBold else FontWeight.Normal, color = when { isSelected -> Color.White; isClickable -> TrainRTheme.NavyDark; else -> TrainRTheme.TextSub.copy(alpha = 0.45f) })
                        }
                    }
                }
            }
        }
    }
}

fun pickFilesDialog(title: String, extensions: Array<String>, multiSelect: Boolean): List<File> {
    val chooser = JFileChooser().apply { dialogTitle = title; isMultiSelectionEnabled = multiSelect; fileFilter = FileNameExtensionFilter("Поддерживаемые файлы (${extensions.joinToString(", ") { "*.$it" }})", *extensions) }
    val res = chooser.showOpenDialog(null)
    if (res == JFileChooser.APPROVE_OPTION) { return if (multiSelect) chooser.selectedFiles?.toList() ?: emptyList() else listOfNotNull(chooser.selectedFile) }
    return emptyList()
}
