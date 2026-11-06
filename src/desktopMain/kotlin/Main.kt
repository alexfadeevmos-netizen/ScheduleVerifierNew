@file:JvmName("MainKt")
package com.scheduleverifier.desktop

import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.Canvas
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.horizontalScroll
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
import androidx.compose.runtime.getValue
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.geometry.CornerRadius
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.geometry.Size
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.Path
import androidx.compose.ui.graphics.drawscope.Stroke
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
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.nio.charset.Charset
import java.time.LocalDate
import java.time.YearMonth
import java.time.format.DateTimeFormatter
import java.util.Locale
import javax.swing.JFileChooser
import javax.swing.filechooser.FileNameExtensionFilter

// ============================================================================
// 1. ЯВНЫЕ ТИПЫ И МОДЕЛИ ДАННЫХ (ЗАЩИТА КОМПИЛЯТОРА K2/FIR ОТ >>>>)
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
    EXPRESS("Экспрессы")
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
// 2. ЦВЕТОВАЯ ПАЛИТРА И КОНСТАНТЫ ПОЛИГОНА ОКТЯБРЬСКОЙ ЖД
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
        "Малино", "Рижская", "Петровско-Разумовская (старая)"
    )

    fun normalizeStationName(raw: String): String {
        val clean = raw.trim()
            .replace("Зеленоград-Крюково", "Крюково", ignoreCase = true)
            .replace("Грачевская", "Грачёвская", ignoreCase = true)
            .replace("Березки-Дачные", "Берёзки-Дачные", ignoreCase = true)
            .replace("Москва-Пасс.-Октябрьская", "Москва", ignoreCase = true)
            .replace("Ленинградский вокзал", "Москва", ignoreCase = true)
        return ALL_STATIONS.firstOrNull { it.equals(clean, ignoreCase = true) } ?: clean
    }

    fun isTrainToMoscow(trainNumber: String): Boolean {
        val firstPart = trainNumber.split("/").firstOrNull()?.trim() ?: trainNumber
        val digits = Regex("\\d+").find(firstPart)?.value ?: return true
        val lastDigit = digits.lastOrNull()?.digitToIntOrNull() ?: return true
        return lastDigit % 2 != 0
    }

    fun isMcdTrain(trainNumber: String, route: String): Boolean {
        if (trainNumber.contains("/")) return true
        val digits = Regex("\\d{4}").find(trainNumber)?.value?.toIntOrNull() ?: 0
        if (digits in 7200..7499) return true
        if (route.contains("Ипподром", ignoreCase = true) || route.contains("47 км", ignoreCase = true)) return true
        return false
    }

    fun isExpressTrain(trainNumber: String, route: String): Boolean {
        if (isMcdTrain(trainNumber, route)) return false
        val digits = Regex("\\d{4}").find(trainNumber)?.value?.toIntOrNull() ?: 0
        return digits in 7000..7999
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
        title = "Train-R • Аналитический портал сверки расписания и генерации отмен (АО «МТ ППК» • МЦД-3)"
    ) {
        MaterialTheme {
            TrainRApp()
        }
    }
}

// ============================================================================
// 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ЛОГИКИ, ПАРСИНГА И АНАЛИТИКИ
// ============================================================================

fun buildRemainingStopsNote(
    skippedStops: List<String>,
    addedStops: List<String>,
    isTechStop: Boolean = false
): String {
    val parts = mutableListOf<String>()
    if (skippedStops.isNotEmpty()) {
        parts.add("без ост.: " + skippedStops.distinct().joinToString(", "))
    }
    if (addedStops.isNotEmpty()) {
        parts.add("назн. ост.: " + addedStops.distinct().joinToString(", "))
    }
    if (isTechStop) {
        parts.add("тех. стоянка (без посадки/высадки пассажиров)")
    }
    return if (parts.isEmpty()) "по графику со всеми остановками" else parts.joinToString("; ")
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
            val latestTelegramTime = sameTrains.map { it.telegramTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: base.telegramTime
            val latestExportTime = sameTrains.map { it.exportTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: base.exportTime
            val latestYandexTime = sameTrains.map { it.yandexTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: base.yandexTime

            val mergedActualText = sameTrains.map { it.actualChangesText.trim() }
                .filter { it.isNotEmpty() }
                .distinct()
                .joinToString("; ")

            val mergedWordLine = sameTrains.map { it.wordExportLine.trim() }
                .filter { it.isNotEmpty() }
                .distinct()
                .joinToString("; ")

            val mergedBasis = sameTrains.map { it.telegramBasis.trim() }
                .filter { it.isNotEmpty() && it != "—" }
                .distinct()
                .joinToString(", ")
                .ifEmpty { base.telegramBasis }

            val mergedError = sameTrains.mapNotNull { it.globalErrorNote }.distinct().joinToString("; ").ifEmpty { null }
            val mergedRaw = sameTrains.map { it.rawTelegramSnippet }.distinct().joinToString("\n---\n")

            val evaluatedStatus = evaluateFourWayAuditStatus(
                afishaTime = base.afishaTime,
                telegramTime = latestTelegramTime,
                exportTime = latestExportTime,
                yandexTime = latestYandexTime,
                isCanceled = combinedCanceled,
                skippedStops = combinedSkipped,
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
                    skippedStops = combinedSkipped,
                    addedStops = combinedAdded,
                    globalErrorNote = mergedError,
                    rawTelegramSnippet = mergedRaw,
                    auditStatus = evaluatedStatus
                )
            )
        }
    }
    return mergedList.sortedBy { row ->
        val t = if (row.telegramTime != "—" && row.telegramTime != "ОТМЕНЕН") row.telegramTime else row.afishaTime
        Regex("\\d{2}:\\d{2}").find(t)?.value ?: "99:99"
    }
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
    if (isCanceled && yandexTime != "—" && yandexTime != "ОТМЕНЕН") {
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
// 5. ЧТЕНИЕ И КЭШИРОВАНИЕ ФАЙЛОВ (UTF-8 / CP1251, EXCEL, WORD, PDF)
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
        val utf8Str = String(bytes, Charsets.UTF_8)
        val cp1251Str = String(bytes, Charset.forName("windows-1251"))

        val cyrillicCountUtf8 = Regex("[А-Яа-яЁё]").findAll(utf8Str).count()
        val cyrillicCount1251 = Regex("[А-Яа-яЁё]").findAll(cp1251Str).count()

        return if ( utf8Str.contains("\uFFFD") || cyrillicCount1251 > cyrillicCountUtf8 ) {
            cp1251Str
        } else {
            utf8Str
        }
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
                                val m = Regex("\\b([67]\\d{3}(?:/[67]\\d{3})?)\\b").find(cellVal)
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
                                    if (cleanTime != null) {
                                        trainStationTimes.getOrPut(trNum) { mutableMapOf() }[station] = cleanTime
                                    }
                                }
                            }
                            for ((trNum, stMap) in trainStationTimes) {
                                val isToMsk = PolygonConfig.isTrainToMoscow(trNum)
                                val routeStr = if (isToMsk) "Крюково/Тверь — Москва" else "Москва — Крюково/Тверь"
                                result[trNum] = AfishaBaseTrain(
                                    trainNumber = trNum,
                                    route = routeStr,
                                    isToMoscow = isToMsk,
                                    stationTimes = stMap
                                )
                            }
                        }
                    }
                }
            }
        } catch (_: Exception) {
        }
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
                } else {
                    cell.numericCellValue.toLong().toString()
                }
            }
            CellType.FORMULA -> {
                try {
                    cell.stringCellValue.trim()
                } catch (_: Exception) {
                    cell.numericCellValue.toInt().toString()
                }
            }
            else -> ""
        }
    }

    fun parseTelegramFiles(
        files: List<File>,
        afishaMap: Map<String, AfishaBaseTrain>
    ): StationDateTrainMap {
        val rawRowsByStationDate = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        val timePairRegex = Regex("([А-Яа-яЁё\\-\\s\\d]+?)\\s+(\\d{1,2}[.:]\\d{2})(?:\\s*[-–—]\\s*(\\d{1,2}[.:]\\d{2}))?")
        val trainRegex = Regex("№?\\s*([67]\\d{3}(?:/[67]\\d{3})?)")
        val dateRegex = Regex("(\\d{1,2}(?:\\s*,\\s*\\d{1,2})*)\\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)", RegexOption.IGNORE_CASE)

        for (file in files) {
            val fullText = readSmartTextFile(file)
            val isClarification = fullText.contains("Внести разъяснение", ignoreCase = true)
            val basisMatch = Regex("(?:телеграмм[аы]|ТГ|№)\\s*([\\d/\\-А-Яа-я]+)", RegexOption.IGNORE_CASE).find(fullText)
            val basisStr = basisMatch?.value ?: file.nameWithoutExtension

            val detectedDates = mutableListOf<String>()
            for (dm in dateRegex.findAll(fullText)) {
                val daysPart = dm.groupValues[1]
                val monthPart = dm.groupValues[2].lowercase(Locale.getDefault())
                val individualDays = daysPart.split(",").map { it.trim() }.filter { it.isNotEmpty() }
                for (d in individualDays) {
                    detectedDates.add("$d $monthPart")
                }
            }
            if (detectedDates.isEmpty()) {
                detectedDates.add("19 октября")
            }

            val lines = fullText.lines()
            for (line in lines) {
                val trMatch = trainRegex.find(line) ?: continue
                val trainNum = trMatch.groupValues[1]
                val isToMsk = PolygonConfig.isTrainToMoscow(trainNum)
                val isCanceled = line.contains("отменяется", ignoreCase = true) || line.contains("отменить", ignoreCase = true)
                val isAppointed = line.contains("назначается", ignoreCase = true) || line.contains("назначить", ignoreCase = true)

                val closedStationError = PolygonConfig.HARD_CLOSED_STATIONS.firstOrNull { closedSt ->
                    line.contains(closedSt, ignoreCase = true) && !line.contains("без остановки", ignoreCase = true)
                }?.let { "ОШИБКА ТГ: Назначена остановка на закрытой станции $it!" }

                val skippedStops = mutableListOf<String>()
                val addedStops = mutableListOf<String>()
                val stationDepTimes = mutableMapOf<String, String>()

                val stopMatches = timePairRegex.findAll(line).toList()
                for ((idx, sm) in stopMatches.withIndex()) {
                    val rawSt = sm.groupValues[1].trim()
                    val stName = PolygonConfig.normalizeStationName(rawSt)
                    if (stName !in PolygonConfig.ALL_STATIONS) continue

                    val t1 = sm.groupValues[2].replace('.', ':').padStart(5, '0')
                    val t2 = sm.groupValues[3].replace('.', ':').let { if (it.isNotEmpty()) it.padStart(5, '0') else "" }

                    val isTech = line.contains("тех. стоянка", ignoreCase = true) ||
                            line.contains("стоянка техническая", ignoreCase = true) ||
                            line.contains("без пассажиров", ignoreCase = true)

                    if (isClarification) {
                        if (line.contains("посадка, высадка пассажиров", ignoreCase = true)) {
                            addedStops.add(stName)
                            stationDepTimes[stName] = if (t2.isNotEmpty()) t2 else t1
                        } else if (line.contains("стоянка техническая", ignoreCase = true)) {
                            skippedStops.add(stName)
                        }
                    } else {
                        if (t2.isNotEmpty()) {
                            if (isTech) {
                                skippedStops.add(stName)
                            } else {
                               // БЕРЁМ СТРОГО ВТОРОЕ ВРЕМЯ (ВРЕМЯ ОТПРАВЛЕНИЯ)
                                stationDepTimes[stName] = t2
                            }
                        } else {
                            // ОДНО ВРЕМЯ: Если начальная станция — это время отправления, если промежуточная — проследование без остановки!
                            if (idx == 0) {
                                stationDepTimes[stName] = t1
                            } else {
                                skippedStops.add(stName)
                            }
                        }
                    }
                }

                val targetStations = if (stationDepTimes.isNotEmpty()) {
                    stationDepTimes.keys
                } else {
                    setOf("Крюково", "Химки", "Сходня", "Клин", "Подсолнечная", "Тверь")
                }

                for (st in targetStations) {
                    // На Левобережной закрыта платформа №2 (путь №2 «от Москвы»)
                    if (st == "Левобережная" && !isToMsk) continue

                    val afishaT = afishaMap[trainNum]?.stationTimes?.get(st) ?: "—"
                    val tgT = if (isCanceled) "ОТМЕНЕН" else (stationDepTimes[st] ?: afishaT)
                    val routeStr = afishaMap[trainNum]?.route ?: if (isToMsk) "Крюково — Москва" else "Москва — Крюково"
                    val isMcd = PolygonConfig.isMcdTrain(trainNum, routeStr)
                    val isExp = PolygonConfig.isExpressTrain(trainNum, routeStr)

                    val remNote = buildRemainingStopsNote(skippedStops, addedStops)
                    val changeSummary = when {
                        isCanceled -> "Поезд №$trainNum отменён на участке следования"
                        isAppointed -> "Назначается поезд №$trainNum отпр. в $tgT ($remNote)"
                        else -> "Поезд №$trainNum отпр. в $tgT ($remNote)"
                    }

                    val wordText = when {
                        isCanceled -> "№$trainNum ($routeStr) отправлением в $afishaT — ОТМЕНЯЕТСЯ"
                        else -> "№$trainNum ($routeStr) отправлением в $afishaT — отправится в $tgT ($remNote)"
                    }

                    for (dateKey in detectedDates.distinct()) {
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
                            actualChangesText = changeSummary,
                            wordExportLine = wordText,
                            isCanceled = isCanceled,
                            isAppointed = isAppointed,
                            skippedStops = skippedStops,
                            addedStops = addedStops,
                            globalErrorNote = closedStationError,
                            rawTelegramSnippet = line.trim(),
                            auditStatus = evaluateFourWayAuditStatus(
                                afishaTime = afishaT,
                                telegramTime = tgT,
                                exportTime = tgT,
                                yandexTime = "—",
                                isCanceled = isCanceled,
                                skippedStops = skippedStops,
                                globalError = closedStationError
                            )
                        )
                        rawRowsByStationDate
                            .getOrPut(st) { mutableMapOf() }
                            .getOrPut(dateKey) { mutableListOf() }
                            .add(rowObj)
                    }
                }
            }
        }

        val finalMap = mutableMapOf<String, Map<String, Map<String, TrainChangeRow>>>()
        for ((st, dateMap) in rawRowsByStationDate) {
            val innerDateMap = mutableMapOf<String, Map<String, TrainChangeRow>>()
            for ((dt, list) in dateMap) {
                val merged = mergeSameTrainRowsForDay(list)
                innerDateMap[dt] = merged.associateBy { it.trainNumber }
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
        val fileNameLower = file.name.lowercase(Locale.getDefault())
        val defaultIsToMoscow = !fileNameLower.contains("от_москвы") && !fileNameLower.contains("от москвы")

        val mutableResult = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        for ((st, dtMap) in existingMap) {
            val mDt = mutableMapOf<String, MutableList<TrainChangeRow>>()
            for ((dt, trMap) in dtMap) {
                mDt[dt] = trMap.values.toMutableList()
            }
            mutableResult[st] = mDt
        }

        var currentStation = "Крюково"
        var currentDate = "19 октября"

        val stationHeaderRegex = Regex("^\\s*([А-Яа-яЁё\\-\\s123]+?)\\s+Изменения\\s+в\\s+графике", RegexOption.IGNORE_CASE)
        val standaloneStationRegex = Regex("^\\s*(?:Станция|ст\\.)?\\s*([А-Яа-яЁё\\-1]+)\\s*$", RegexOption.IGNORE_CASE)
        val dateLineRegex = Regex("^\\s*(\\d{1,2}(?:\\s*,\\s*\\d{1,2})*)\\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*$", RegexOption.IGNORE_CASE)
        val arrivalFilterRegex = Regex("(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)", RegexOption.IGNORE_CASE)
        val trainNumberRegex = Regex("№\\s*([67]\\d{3}(?:/[67]\\d{3})?)")

        for (rawLine in text.lines()) {
            val trimmed = rawLine.trim()
            if (trimmed.isEmpty()) continue

            // Строго исключаем строки с «прибытием в / прибывает в / прибудет в»
            if (arrivalFilterRegex.containsMatchIn(trimmed)) continue

            val stHeaderMatch = stationHeaderRegex.find(trimmed)
            if (stHeaderMatch != null) {
                val candidate = PolygonConfig.normalizeStationName(stHeaderMatch.groupValues[1])
                if (candidate in PolygonConfig.ALL_STATIONS) {
                    currentStation = candidate
                    continue
                }
            }

            val stSoloMatch = standaloneStationRegex.find(trimmed)
            if (stSoloMatch != null) {
                val candidate = PolygonConfig.normalizeStationName(stSoloMatch.groupValues[1])
                if (candidate in PolygonConfig.ALL_STATIONS) {
                    currentStation = candidate
                    continue
                }
            }

            val dtMatch = dateLineRegex.find(trimmed)
            if (dtMatch != null) {
                val daysList = dtMatch.groupValues[1].split(",").map { it.trim() }.filter { it.isNotEmpty() }
                val monthName = dtMatch.groupValues[2].lowercase(Locale.getDefault())
                if (daysList.isNotEmpty()) {
                    currentDate = "${daysList.first()} $monthName"
                }
                continue
            }

            val trMatch = trainNumberRegex.find(trimmed)
            if (trMatch != null) {
                val trNum = trMatch.groupValues[1]
                val isToMsk = if (trimmed.contains("из Москвы", ignoreCase = true) || trimmed.contains("от Москвы", ignoreCase = true)) {
                    false
                } else if (trimmed.contains("до Москвы", ignoreCase = true) || trimmed.contains("на Москву", ignoreCase = true)) {
                    true
                } else {
                    PolygonConfig.isTrainToMoscow(trNum).takeIf { trNum.isNotEmpty() } ?: defaultIsToMoscow
                }

                // Исключаем Левобережную от Москвы (платформа №2 закрыта)
                if (currentStation == "Левобережная" && !isToMsk) continue

                val timesFound = Regex("(\\d{1,2}:\\d{2})").findAll(trimmed).map { it.groupValues[1].padStart(5, '0') }.toList()
                val afishaT = timesFound.firstOrNull() ?: (afishaMap[trNum]?.stationTimes?.get(currentStation) ?: "—")
                val isCanceled = trimmed.contains("отменяется", ignoreCase = true) || trimmed.contains("отменен", ignoreCase = true)
                val isAppointed = trimmed.contains("назначается", ignoreCase = true)
                val expT = when {
                    isCanceled -> "ОТМЕНЕН"
                    timesFound.size >= 2 -> timesFound.last()
                    timesFound.size == 1 -> timesFound.first()
                    else -> "—"
                }

                val routeStr = when {
                    trimmed.contains("Ипподром", ignoreCase = true) -> if (isToMsk) "Крюково — Ипподром" else "Ипподром — Крюково"
                    else -> if (isToMsk) "до Москвы" else "из Москвы"
                }

                val row = TrainChangeRow(
                    id = "${currentDate}_${currentStation}_${trNum}_${System.nanoTime()}",
                    dateStr = currentDate,
                    station = currentStation,
                    trainNumber = trNum,
                    isToMoscow = isToMsk,
                    isMcd = PolygonConfig.isMcdTrain(trNum, trimmed),
                    isExpress = PolygonConfig.isExpressTrain(trNum, trimmed),
                    route = routeStr,
                    telegramBasis = "Выгрузка (${file.name})",
                    afishaTime = afishaT,
                    telegramTime = expT,
                    exportTime = expT,
                    yandexTime = "—",
                    actualChangesText = trimmed,
                    wordExportLine = trimmed,
                    isCanceled = isCanceled,
                    isAppointed = isAppointed,
                    rawTelegramSnippet = trimmed
                )

                mutableResult
                    .getOrPut(currentStation) { mutableMapOf() }
                    .getOrPut(currentDate) { mutableListOf() }
                    .add(row)
            }
        }

        val out = mutableMapOf<String, Map<String, Map<String, TrainChangeRow>>>()
        for ((st, dtMap) in mutableResult) {
            val dOut = mutableMapOf<String, Map<String, TrainChangeRow>>()
            for ((dt, rows) in dtMap) {
                dOut[dt] = mergeSameTrainRowsForDay(rows).associateBy { it.trainNumber }
            }
            out[st] = dOut
        }
        return out
    }

    fun parseYandexPdf(file: File): List<YandexParsedTrain> {
        val text = readSmartTextFile(file)
        val parsed = mutableListOf<YandexParsedTrain>()
        val lineRegex = Regex("([67]\\d{3}(?:/[67]\\d{3})?).*?(\\d{1,2}:\\d{2})")
        for (line in text.lines()) {
            val m = lineRegex.find(line) ?: continue
            val trNum = m.groupValues[1]
            val depTime = m.groupValues[2].padStart(5, '0')
            parsed.add(
                YandexParsedTrain(
                    trainNumber = trNum,
                    dateStr = "19 октября",
                    station = "Крюково",
                    departureTime = depTime,
                    statusText = if (line.contains("отмен", ignoreCase = true)) "Отменён" else "По расписанию"
                )
            )
        }
        return parsed
    }
}

// ============================================================================
// 6. ГЕНЕРАТОР ФАЙЛОВ WORD (.DOCX) ПО ЭТАЛОНУ «ОТМЕНЫ ПО СТАНЦИЯМ»
// ============================================================================

object WordCancelGenerator {
    private val arrivalExcludeRegex = Regex("(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)", RegexOption.IGNORE_CASE)
    private val nightNoteRegex = Regex("\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)", RegexOption.IGNORE_CASE)

    fun getDirectionSubtitle(station: String, isToMoscow: Boolean): String {
        return if (isToMoscow) {
            if (station in PolygonConfig.MCD3_STATIONS) {
                "на Ленинградский вокзал, Ипподром"
            } else {
                "на Москву"
            }
        } else {
            "от Москвы"
        }
    }

    fun generateStationWordFiles(
        selectedStations: Set<String>,
        selectedDates: Set<String>,
        directionTab: DirectionTab,
        allData: StationDateTrainMap
    ): Pair<Int, File> {
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

            // Группировка идентичных дат (со стиранием "(в ночь на ...)" для сравнения)
            val signatureToDates = linkedMapOf<String, MutableList<String>>()
            val dateToCleanLines = mutableMapOf<String, List<String>>()

            for (dateStr in datesToProcess) {
                val trainMap = stationDateMap[dateStr] ?: continue
                val filteredRows = trainMap.values
                    .filter { it.isToMoscow == isToMoscow }
                    .filter { !arrivalExcludeRegex.containsMatchIn(it.wordExportLine) && !arrivalExcludeRegex.containsMatchIn(it.actualChangesText) }

                if (filteredRows.isEmpty()) continue

                val rawLines = filteredRows.map { it.wordExportLine.trim() }
                dateToCleanLines[dateStr] = rawLines

                val normalizedSignature = rawLines.joinToString("\n") { line ->
                    nightNoteRegex.replace(line, "").trim()
                }
                signatureToDates.getOrPut(normalizedSignature) { mutableListOf() }.add(dateStr)
            }

            if (signatureToDates.isEmpty()) continue

            val fileNameSuffix = if (isToMoscow) "на Москву" else "от Москвы"
            val outFile = File(targetFolder, "$station $fileNameSuffix.docx")

            XWPFDocument().use { doc ->
                val subtitle = getDirectionSubtitle(station, isToMoscow)

                // 1-я строка заголовка (14 pt, по центру, полужирный)
                addCenteredBoldHeader(doc, "$station  Изменения в графике движения эл/поездов")
                // 2-я строка подзаголовка направления (14 pt, по центру, полужирный)
                addCenteredBoldHeader(doc, subtitle)

                var currentPageLineCount = 2

                for ((_, groupedDates) in signatureToDates) {
                    val isMergedMultipleDays = groupedDates.size > 1
                    val formattedDatesHeader = formatMergedDatesHeader(groupedDates)

                    if (currentPageLineCount >= 47) {
                        addContinuationHeader(doc, station, formattedDatesHeader, subtitle)
                        currentPageLineCount = 2
                    }

                    // Строка даты (по центру, полужирный, 14 pt)
                    addCenteredBoldHeader(doc, formattedDatesHeader)
                    currentPageLineCount++

                    val sampleDate = groupedDates.first()
                    val trainLines = dateToCleanLines[sampleDate] ?: emptyList()

                    for (rawLine in trainLines) {
                        if (currentPageLineCount >= 49) {
                            addContinuationHeader(doc, station, formattedDatesHeader, subtitle)
                            currentPageLineCount = 2
                        }

                        val finalTrainLine = if (isMergedMultipleDays) {
                            nightNoteRegex.replace(rawLine, "").trim()
                        } else {
                            rawLine.trim()
                        }

                        // Строка поезда: 12 pt, по левому краю, отступ 8 пробелов
                        val p = doc.createParagraph()
                        p.alignment = ParagraphAlignment.LEFT
                        p.spacingAfter = 0
                        p.spacingBefore = 0
                        val run = p.createRun()
                        run.fontFamily = "Times New Roman"
                        run.fontSize = 12
                        run.isBold = false
                        run.setText("        $finalTrainLine")

                        currentPageLineCount++
                    }
                }

                FileOutputStream(outFile).use { fos ->
                    doc.write(fos)
                }
                generatedCount++
            }
        }

        return Pair(generatedCount, targetFolder)
    }

    private fun addCenteredBoldHeader(doc: XWPFDocument, text: String) {
        val p = doc.createParagraph()
        p.alignment = ParagraphAlignment.CENTER
        p.spacingAfter = 40
        p.spacingBefore = 40
        val run = p.createRun()
        run.fontFamily = "Times New Roman"
        run.fontSize = 14
        run.isBold = true
        run.setText(text)
    }

    private fun addContinuationHeader(
        doc: XWPFDocument,
        station: String,
        datesHeader: String,
        subtitle: String
    ) {
        val p1 = doc.createParagraph()
        p1.isPageBreak = true
        p1.alignment = ParagraphAlignment.CENTER
        val r1 = p1.createRun()
        r1.fontFamily = "Times New Roman"
        r1.fontSize = 14
        r1.isBold = true
        r1.setText("$station продолжение на $datesHeader")

        val p2 = doc.createParagraph()
        p2.alignment = ParagraphAlignment.CENTER
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
            val dayNum = parts.firstOrNull() ?: d
            val monthStr = parts.drop(1).joinToString(" ")
            Pair(dayNum, monthStr)
        }
        val byMonth = parsed.groupBy({ it.second }, { it.first })
        return byMonth.entries.joinToString(", ") { (month, days) ->
            val sortedDays = days.sortedBy { it.toIntOrNull() ?: 99 }
            if (month.isNotBlank()) "${sortedDays.joinToString(", ")} $month" else sortedDays.joinToString(", ")
        }
    }
}

// ============================================================================
// 7. ДЕМО-ДАННЫЕ ПО УМОЛЧАНИЮ ДЛЯ МГНОВЕННОЙ ПРОВЕРКИ РАБОТЫ
// ============================================================================

fun createInitialDemoSchedule(): StationDateTrainMap {
    val sampleRowsKryukovo = listOf(
        TrainChangeRow(
            id = "demo_1",
            dateStr = "19 октября",
            station = "Крюково",
            trainNumber = "7529",
            isToMoscow = true,
            isMcd = false,
            isExpress = true,
            route = "Крюково — Москва",
            telegramBasis = "ТГ № 1412/ОКТ от 14.10.2026",
            afishaTime = "04:12",
            telegramTime = "04:36",
            exportTime = "04:36",
            yandexTime = "04:36",
            actualChangesText = "Отпр. 04:36 вместо 04:12; без ост.: Сходня",
            wordExportLine = "№7529 до Москвы отправлением в 04:12 - отправится в 04:36 (в ночь на 20 октября)",
            skippedStops = listOf("Сходня"),
            rawTelegramSnippet = "19, 20, 21, 22 октября поезд №7529 Крюково 04:36 Сходня 04:43 Химки 04:51-04:52 Москва 05:14",
            fullThread = listOf(
                StationStopEntry("Крюково", "04:12", "04:36", "04:36", "04:36"),
                StationStopEntry("Сходня", "04:19", "04:43 (прослед.)", "—", "—", isSkippedByTelegram = true),
                StationStopEntry("Химки", "04:28", "04:52", "04:52", "04:52"),
                StationStopEntry("Москва", "04:50", "05:14", "05:14", "05:14")
            ),
            auditStatus = AuditStatus.OK
        ),
        TrainChangeRow(
            id = "demo_2",
            dateStr = "19 октября",
            station = "Крюково",
            trainNumber = "7301/7302",
            isToMoscow = true,
            isMcd = true,
            isExpress = false,
            route = "Крюково — Ипподром",
            telegramBasis = "ТГ № 1415/ОКТ от 15.10.2026",
            afishaTime = "05:25",
            telegramTime = "05:35",
            exportTime = "05:35",
            yandexTime = "05:25",
            actualChangesText = "МЦД-3: сдвиг отправления на +10 мин (в Яндексе старое время 05:25!)",
            wordExportLine = "№7301/7302 на Ипподром отправлением в 05:25 - отправится в 05:35",
            rawTelegramSnippet = "19 октября п. №7301/7302 Крюково 05:35 Фирсановская 05:39-05:40 Сходня 05:43-05:44",
            fullThread = listOf(
                StationStopEntry("Крюково", "05:25", "05:35", "05:35", "05:25"),
                StationStopEntry("Фирсановская", "05:29", "05:40", "05:40", "05:29"),
                StationStopEntry("Сходня", "05:33", "05:44", "05:44", "05:33"),
                StationStopEntry("Химки", "05:48", "05:58", "05:58", "05:48")
            ),
            auditStatus = AuditStatus.TIME_MISMATCH
        ),
        TrainChangeRow(
            id = "demo_3",
            dateStr = "19 октября",
            station = "Крюково",
            trainNumber = "6700",
            isToMoscow = false,
            isMcd = false,
            isExpress = false,
            route = "Москва — Тверь",
            telegramBasis = "ТГ № 1418/ОКТ от 15.10.2026",
            afishaTime = "06:14",
            telegramTime = "06:18",
            exportTime = "06:18",
            yandexTime = "06:18",
            actualChangesText = "Отпр. 06:18 вместо 06:14; без ост.: Алабушево, Покровка",
            wordExportLine = "№6700 из Москвы до Твери отправлением в 06:14 - отправится в 06:18 (без ост.: Алабушево, Покровка)",
            skippedStops = listOf("Алабушево", "Покровка"),
            rawTelegramSnippet = "19 октября п. №6700 Крюково 06:17-06:18 Алабушево 06:23 Подсолнечная 06:40-06:41 Покровка 06:52 Клин 07:02-07:03",
            fullThread = listOf(
                StationStopEntry("Москва", "05:22", "05:26", "05:26", "05:26"),
                StationStopEntry("Химки", "05:47", "05:51", "05:51", "05:51"),
                StationStopEntry("Крюково", "06:14", "06:18", "06:18", "06:18"),
                StationStopEntry("Алабушево", "06:19", "06:23 (прослед.)", "—", "—", isSkippedByTelegram = true),
                StationStopEntry("Подсолнечная", "06:37", "06:41", "06:41", "06:41"),
                StationStopEntry("Покровка", "06:48", "06:52 (прослед.)", "—", "—", isSkippedByTelegram = true),
                StationStopEntry("Клин", "06:59", "07:03", "07:03", "07:03")
            ),
            auditStatus = AuditStatus.OK
        ),
        TrainChangeRow(
            id = "demo_4",
            dateStr = "20 октября",
            station = "Крюково",
            trainNumber = "7529",
            isToMoscow = true,
            isMcd = false,
            isExpress = true,
            route = "Крюково — Москва",
            telegramBasis = "ТГ № 1412/ОКТ от 14.10.2026",
            afishaTime = "04:12",
            telegramTime = "04:36",
            exportTime = "04:36",
            yandexTime = "04:36",
            actualChangesText = "Отпр. 04:36 вместо 04:12; без ост.: Сходня",
            wordExportLine = "№7529 до Москвы отправлением в 04:12 - отправится в 04:36 (в ночь на 21 октября)",
            skippedStops = listOf("Сходня"),
            rawTelegramSnippet = "19, 20, 21, 22 октября поезд №7529 Крюково 04:36 Сходня 04:43 Химки 04:51-04:52 Москва 05:14",
            auditStatus = AuditStatus.OK
        )
    )

    val dateMap = mapOf(
        "19 октября" to sampleRowsKryukovo.filter { it.dateStr == "19 октября" }.associateBy { it.trainNumber },
        "20 октября" to sampleRowsKryukovo.filter { it.dateStr == "20 октября" }.associateBy { it.trainNumber }
    )
    return mapOf("Крюково" to dateMap)
}

// ============================================================================
// 8. UI КОМПОНЕНТЫ (ПОЛНЫЙ КОМПЛЕКТ СОГЛАСНО ТЕХНИЧЕСКОМУ ЗАДАНИЮ)
// ============================================================================

@OptIn(ExperimentalLayoutApi::class)
@Composable
fun TrainRApp() {
    val scope = rememberCoroutineScope()

    var currentNavTab: MainNavTab by remember { mutableStateOf(MainNavTab.UPLOADS) }
    var currentDirectionTab: DirectionTab by remember { mutableStateOf(DirectionTab.TO_MOSCOW) }
    var currentTrainFilter: TrainFilterType by remember { mutableStateOf(TrainFilterType.ALL) }

    var searchQuery: String by remember { mutableStateOf("") }
    var systemStatusText: String by remember { mutableStateOf("Система готова. Полигон: Октябрьская ЖД / МЦД-3") }

    var afishaCacheMap: Map<String, AfishaBaseTrain> by remember { mutableStateOf(emptyMap()) }
    var stationScheduleMap: StationDateTrainMap by remember { mutableStateOf(createInitialDemoSchedule()) }
    var yandexTrainsList: List<YandexParsedTrain> by remember { mutableStateOf(emptyList()) }

    var loadedYandexCount: Int by remember { mutableStateOf(0) }
    var loadedAfishaCount: Int by remember { mutableStateOf(0) }
    var loadedTelegramsCount: Int by remember { mutableStateOf(3) }
    var loadedExportsCount: Int by remember { mutableStateOf(1) }

    var activePreviewStation: String by remember { mutableStateOf("Крюково") }
    var checkedStationsForWord: Set<String> by remember { mutableStateOf(setOf("Крюково", "Химки", "Сходня")) }
    var selectedFilterDates: Set<String> by remember { mutableStateOf(emptySet()) }

    var showCalendarDialog: Boolean by remember { mutableStateOf(false) }
    var selectedTrainForDossier: TrainChangeRow? by remember { mutableStateOf(null) }

    // Автозагрузка кэшированной Афиши при старте приложения
    LaunchedEffect(Unit) {
        if (FileParsers.cachedAfishaFile.exists()) {
            withContext(Dispatchers.IO) {
                val parsed = FileParsers.parseAfishaExcel(FileParsers.cachedAfishaFile)
                if (parsed.isNotEmpty()) {
                    afishaCacheMap = parsed
                    loadedAfishaCount = parsed.size
                    systemStatusText = "Загружен кэш базовой Афиши (${parsed.size} поездов)"
                }
            }
        }
    }

    val allDatesWithChanges: Set<String> = remember(stationScheduleMap) {
        stationScheduleMap.values.flatMap { it.keys }.toSet()
    }

    val currentTableRows: List<TrainChangeRow> = remember(
        stationScheduleMap,
        activePreviewStation,
        selectedFilterDates,
        currentDirectionTab,
        currentTrainFilter,
        searchQuery
    ) {
        val dateMap = stationScheduleMap[activePreviewStation] ?: emptyMap()
        val datesToInclude = if (selectedFilterDates.isEmpty()) dateMap.keys else dateMap.keys.intersect(selectedFilterDates)
        val rawList = datesToInclude.flatMap { dt -> dateMap[dt]?.values ?: emptyList() }

        rawList
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

    Row(modifier = Modifier.fillMaxSize().background(TrainRTheme.BgLight)) {
        // ЛЕВАЯ НАВИГАЦИОННАЯ ПАНЕЛЬ (Sidebar 220.dp)
        Column(
            modifier = Modifier
                .width(220.dp)
                .fillMaxHeight()
                .background(TrainRTheme.SidebarNavy)
                .padding(16.dp),
            verticalArrangement = Arrangement.SpaceBetween
        ) {
            Column {
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    modifier = Modifier.padding(bottom = 24.dp)
                ) {
                    VintageTrainLogo(modifier = Modifier.size(44.dp))
                    Spacer(modifier = Modifier.width(12.dp))
                    Column {
                        Text(
                            text = "Train-R",
                            color = Color.White,
                            fontSize = 20.sp,
                            fontWeight = FontWeight.ExtraBold
                        )
                        Text(
                            text = "МТППК • МЦД-3",
                            color = TrainRTheme.GoldAccent,
                            fontSize = 11.sp,
                            fontWeight = FontWeight.SemiBold
                        )
                    }
                }

                HorizontalDivider(color = Color.White.copy(alpha = 0.12f))
                Spacer(modifier = Modifier.height(16.dp))

                SidebarNavButton(
                    title = MainNavTab.UPLOADS.title,
                    icon = Icons.Default.Description,
                    isSelected = currentNavTab == MainNavTab.UPLOADS,
                    onClick = { currentNavTab = MainNavTab.UPLOADS }
                )
                Spacer(modifier = Modifier.height(8.dp))
                SidebarNavButton(
                    title = MainNavTab.REPORTS.title,
                    icon = Icons.Default.Assessment,
                    isSelected = currentNavTab == MainNavTab.REPORTS,
                    onClick = { currentNavTab = MainNavTab.REPORTS }
                )
                Spacer(modifier = Modifier.height(8.dp))
                SidebarNavButton(
                    title = MainNavTab.YANDEX.title,
                    icon = Icons.Default.FactCheck,
                    isSelected = currentNavTab == MainNavTab.YANDEX,
                    onClick = { currentNavTab = MainNavTab.YANDEX }
                )
            }

            // Блок «Статус системы» внизу сайдбара
            Surface(
                color = TrainRTheme.SidebarActive,
                shape = RoundedCornerShape(10.dp),
                border = BorderStroke(1.dp, Color.White.copy(alpha = 0.1f)),
                modifier = Modifier.fillMaxWidth()
            ) {
                Column(modifier = Modifier.padding(12.dp)) {
                    Row(verticalAlignment = Alignment.CenterVertically) {
                        Box(
                            modifier = Modifier
                                .size(8.dp)
                                .clip(CircleShape)
                                .background(TrainRTheme.GreenSuccess)
                        )
                        Spacer(modifier = Modifier.width(8.dp))
                        Text(
                            text = "СТАТУС СИСТЕМЫ",
                            color = Color.White.copy(alpha = 0.7f),
                            fontSize = 10.sp,
                            fontWeight = FontWeight.Bold
                        )
                    }
                    Spacer(modifier = Modifier.height(6.dp))
                    Text(
                        text = systemStatusText,
                        color = Color.White,
                        fontSize = 11.sp,
                        lineHeight = 14.sp
                    )
                }
            }
        }

        // ОСНОВНАЯ РАБОЧАЯ ОБЛАСТЬ СПРАВА
        Column(modifier = Modifier.weight(1f).fillMaxHeight()) {
            // ВЕРХНЯЯ ПАНЕЛЬ (Top Bar)
            Surface(
                color = Color.White,
                shadowElevation = 2.dp,
                modifier = Modifier.fillMaxWidth()
            ) {
                Row(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 20.dp, vertical = 12.dp),
                    verticalAlignment = Alignment.CenterVertically,
                    horizontalArrangement = Arrangement.SpaceBetween
                ) {
                    // Поле быстрого поиска по номеру поезда
                    OutlinedTextField(
                        value = searchQuery,
                        onValueChange = { searchQuery = it },
                        placeholder = {
                            Text(
                                text = "Поиск поезда (№6504, 7529...)",
                                fontSize = 13.sp,
                                color = TrainRTheme.TextSub
                            )
                        },
                        leadingIcon = {
                            Icon(Icons.Default.Search, contentDescription = null, tint = TrainRTheme.TextSub)
                        },
                        trailingIcon = {
                            if (searchQuery.isNotEmpty()) {
                                IconButton(onClick = { searchQuery = "" }) {
                                    Icon(Icons.Default.Clear, contentDescription = "Очистить")
                                }
                            }
                        },
                        singleLine = true,
                        shape = RoundedCornerShape(8.dp),
                        modifier = Modifier.width(320.dp).height(50.dp)
                    )

                    // 4 кнопки загрузки источников данных
                    Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                        TopBarLoadButton(
                            label = "ЯНДЕКС",
                            count = loadedYandexCount,
                            badgeColor = TrainRTheme.YandexRed,
                            onClick = {
                                val file = pickFilesDialog(
                                    title = "Выберите PDF расписание Яндекса",
                                    extensions = arrayOf("pdf"),
                                    multiSelect = false
                                ).firstOrNull()
                                if (file != null) {
                                    scope.launch(Dispatchers.IO) {
                                        val parsed = FileParsers.parseYandexPdf(file)
                                        yandexTrainsList = parsed
                                        loadedYandexCount = parsed.size
                                        systemStatusText = "Загружен Яндекс PDF: ${file.name} (${parsed.size} записей)"
                                    }
                                }
                            }
                        )

                        TopBarLoadButton(
                            label = "АФИША",
                            count = loadedAfishaCount,
                            badgeColor = TrainRTheme.RoyalBlue,
                            onClick = {
                                val file = pickFilesDialog(
                                    title = "Выберите базовое расписание Афиши (.xlsx)",
                                    extensions = arrayOf("xlsx", "xls"),
                                    multiSelect = false
                                ).firstOrNull()
                                if (file != null) {
                                    scope.launch(Dispatchers.IO) {
                                        val cached = FileParsers.ensureCacheAfisha(file)
                                        val parsed = FileParsers.parseAfishaExcel(cached)
                                        afishaCacheMap = parsed
                                        loadedAfishaCount = parsed.size
                                        systemStatusText = "Афиша закэширована: ${file.name} (${parsed.size} поездов)"
                                    }
                                }
                            }
                        )

                        TopBarLoadButton(
                            label = "ТЕЛЕГРАММА",
                            count = loadedTelegramsCount,
                            badgeColor = TrainRTheme.AmberWarn,
                            onClick = {
                                val files = pickFilesDialog(
                                    title = "Выберите оперативные телеграммы и разъяснения (.docx, .pdf, .txt)",
                                    extensions = arrayOf("docx", "pdf", "txt"),
                                    multiSelect = true
                                )
                                if (files.isNotEmpty()) {
                                    scope.launch(Dispatchers.IO) {
                                        val parsed = FileParsers.parseTelegramFiles(files, afishaCacheMap)
                                        stationScheduleMap = parsed
                                        loadedTelegramsCount = files.size
                                        systemStatusText = "Обработано телеграмм: ${files.size} шт."
                                    }
                                }
                            }
                        )

                        TopBarLoadButton(
                            label = "ВЫГРУЗКА",
                            count = loadedExportsCount,
                            badgeColor = TrainRTheme.GreenSuccess,
                            onClick = {
                                val files = pickFilesDialog(
                                    title = "Выберите сводные файлы выгрузок (.txt, .docx)",
                                    extensions = arrayOf("txt", "docx"),
                                    multiSelect = true
                                )
                                if (files.isNotEmpty()) {
                                    scope.launch(Dispatchers.IO) {
                                        var updatedMap = stationScheduleMap
                                        for (f in files) {
                                            updatedMap = FileParsers.parseSummaryExportFile(f, afishaCacheMap, updatedMap)
                                        }
                                        stationScheduleMap = updatedMap
                                        loadedExportsCount = files.size
                                        systemStatusText = "Загружена сводная выгрузка (${files.size} ф.)"
                                    }
                                }
                            }
                        )
                    }
                }
            }

            // ПАНЕЛЬ УПРАВЛЕНИЯ И ВЫБОРА СТАНЦИЙ (под верхней панелью)
            Surface(
                color = Color.White,
                border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                modifier = Modifier.fillMaxWidth().padding(horizontal = 20.dp, vertical = 10.dp),
                shape = RoundedCornerShape(10.dp)
            ) {
                Column(modifier = Modifier.padding(14.dp)) {
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.SpaceBetween,
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Row(horizontalArrangement = Arrangement.spacedBy(10.dp), verticalAlignment = Alignment.CenterVertically) {
                            OutlinedButton(
                                onClick = { showCalendarDialog = true },
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.DateRange, contentDescription = null, modifier = Modifier.size(18.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                val dateBtnLabel = if (selectedFilterDates.isEmpty()) {
                                    "ФИЛЬТР ПО ДАТЕ (ВСЕ: ${allDatesWithChanges.size})"
                                } else {
                                    "ДАТЫ: ${selectedFilterDates.joinToString(", ")}"
                                }
                                Text(text = dateBtnLabel, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                            }

                            OutlinedButton(
                                onClick = {
                                    systemStatusText = "Выполнена сквозная сверка 4 источников по всем ${PolygonConfig.ALL_STATIONS.size} станциям."
                                    currentNavTab = MainNavTab.REPORTS
                                },
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Sync, contentDescription = null, modifier = Modifier.size(18.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("СВЕРИТЬ ВСЕ", fontWeight = FontWeight.Bold, fontSize = 12.sp)
                            }

                            if (selectedFilterDates.isNotEmpty()) {
                                TextButton(onClick = { selectedFilterDates = emptySet() }) {
                                    Text("Сбросить фильтр дат", fontSize = 12.sp, color = TrainRTheme.RedCancel)
                                }
                            }
                        }

                        // Главная синяя кнопка «СОЗДАТЬ ОТМЕНЫ» (стиль CREATE EVENT)
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
                            contentPadding = PaddingValues(horizontal = 20.dp, vertical = 12.dp)
                        ) {
                            Icon(Icons.Default.Description, contentDescription = null, tint = Color.White)
                            Spacer(modifier = Modifier.width(8.dp))
                            Text(
                                text = "СОЗДАТЬ ОТМЕНЫ (${checkedStationsForWord.size} СТ.)",
                                color = Color.White,
                                fontWeight = FontWeight.ExtraBold,
                                fontSize = 13.sp
                            )
                        }
                    }

                    Spacer(modifier = Modifier.height(10.dp))
                    HorizontalDivider(color = TrainRTheme.BorderCol)
                    Spacer(modifier = Modifier.height(8.dp))

                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.SpaceBetween,
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Text(
                            text = "Станции полигона (чекбокс — генерация Word, клик по названию — предпросмотр):",
                            fontSize = 11.sp,
                            color = TrainRTheme.TextSub,
                            fontWeight = FontWeight.SemiBold
                        )
                        Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                            TextButton(
                                onClick = { checkedStationsForWord = PolygonConfig.ALL_STATIONS.toSet() },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)
                            ) {
                                Text("Выбрать все", fontSize = 11.sp)
                            }
                            TextButton(
                                onClick = { checkedStationsForWord = PolygonConfig.MCD3_STATIONS.intersect(PolygonConfig.ALL_STATIONS.toSet()) },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)
                            ) {
                                Text("Только МЦД-3", fontSize = 11.sp)
                            }
                            TextButton(
                                onClick = { checkedStationsForWord = emptySet() },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)
                            ) {
                                Text("Снять все", fontSize = 11.sp)
                            }
                        }
                    }

                    // FlowRow со списком всех станций и чекбоксами
                    FlowRow(
                        modifier = Modifier.fillMaxWidth().padding(top = 4.dp),
                        horizontalArrangement = Arrangement.spacedBy(6.dp),
                        verticalArrangement = Arrangement.spacedBy(6.dp)
                    ) {
                        for (st in PolygonConfig.ALL_STATIONS) {
                            val isChecked = st in checkedStationsForWord
                            val isPreview = st == activePreviewStation
                            val hasChanges = (stationScheduleMap[st]?.isNotEmpty() == true)

                            Surface(
                                shape = RoundedCornerShape(6.dp),
                                color = when {
                                    isPreview -> TrainRTheme.ActionBlue.copy(alpha = 0.14f)
                                    hasChanges -> Color(0xFFF1F5F9)
                                    else -> Color.White
                                },
                                border = BorderStroke(
                                    width = if (isPreview) 1.5.dp else 1.dp,
                                    color = if (isPreview) TrainRTheme.ActionBlue else TrainRTheme.BorderCol
                                )
                            ) {
                                Row(
                                    verticalAlignment = Alignment.CenterVertically,
                                    modifier = Modifier
                                        .clickable { activePreviewStation = st }
                                        .padding(end = 8.dp, top = 2.dp, bottom = 2.dp)
                                ) {
                                    Checkbox(
                                        checked = isChecked,
                                        onCheckedChange = { checked ->
                                            checkedStationsForWord = if (checked) {
                                                checkedStationsForWord + st
                                            } else {
                                                checkedStationsForWord - st
                                            }
                                        },
                                        modifier = Modifier.size(26.dp)
                                    )
                                    Text(
                                        text = st,
                                        fontSize = 11.sp,
                                        fontWeight = if (isPreview || hasChanges) FontWeight.Bold else FontWeight.Normal,
                                        color = if (isPreview) TrainRTheme.RoyalBlue else TrainRTheme.TextDark
                                    )
                                }
                            }
                        }
                    }
                }
            }

            // ПЕРЕКЛЮЧЕНИЕ МЕЖДУ 3 ОСНОВНЫМИ ВКЛАДКАМИ
            Box(modifier = Modifier.weight(1f).fillMaxWidth().padding(horizontal = 20.dp, vertical = 6.dp)) {
                when (currentNavTab) {
                    MainNavTab.UPLOADS -> {
                        TrainTableSection(
                            activeStation = activePreviewStation,
                            directionTab = currentDirectionTab,
                            onDirectionChange = { currentDirectionTab = it },
                            trainFilter = currentTrainFilter,
                            onTrainFilterChange = { currentTrainFilter = it },
                            rows = currentTableRows,
                            onRowClick = { selectedTrainForDossier = it }
                        )
                    }
                    MainNavTab.REPORTS -> {
                        StationAnalyticsReportView(
                            activeStation = activePreviewStation,
                            rows = currentTableRows,
                            onRowClick = { selectedTrainForDossier = it }
                        )
                    }
                    MainNavTab.YANDEX -> {
                        YandexAuditTabView(
                            activeStation = activePreviewStation,
                            rows = currentTableRows,
                            yandexList = yandexTrainsList,
                            onRowClick = { selectedTrainForDossier = it }
                        )
                    }
                }
            }
        }
    }

    // МОДАЛЬНОЕ ОКНО КАЛЕНДАРЯ (на 2 месяца вперёд)
    if (showCalendarDialog) {
        CalendarGridPickerDialog(
            availableDatesWithChanges = allDatesWithChanges,
            selectedDates = selectedFilterDates,
            onToggleDate = { dt ->
                selectedFilterDates = if (dt in selectedFilterDates) {
                    selectedFilterDates - dt
                } else {
                    selectedFilterDates + dt
                }
            },
            onClearAll = { selectedFilterDates = emptySet() },
            onDismiss = { showCalendarDialog = false }
        )
    }

    // МОДАЛЬНОЕ ОКНО «ДОСЬЕ ПОЕЗДА»
    val dossierTarget = selectedTrainForDossier
    if (dossierTarget != null) {
        TrainDossierDialog(
            trainRow = dossierTarget,
            onDismiss = { selectedTrainForDossier = null }
        )
    }
}

// ============================================================================
// 9. ОТДЕЛЬНЫЕ @Composable КОМПОНЕНТЫ ИНТЕРФЕЙСА
// ============================================================================

@Composable
fun VintageTrainLogo(modifier: Modifier = Modifier) {
    Canvas(modifier = modifier) {
        val w = size.width
        val h = size.height
        drawRoundRect(
            color = TrainRTheme.ActionBlue,
            topLeft = Offset(0f, 0f),
            size = Size(w, h),
            cornerRadius = CornerRadius(12f, 12f)
        )
        // Корпус электропоезда
        val bodyPath = Path().apply {
            moveTo(w * 0.22f, h * 0.20f)
            lineTo(w * 0.78f, h * 0.20f)
            lineTo(w * 0.85f, h * 0.72f)
            lineTo(w * 0.15f, h * 0.72f)
            close()
        }
        drawPath(path = bodyPath, color = Color.White)
        // Лобовое стекло
        drawRoundRect(
            color = TrainRTheme.NavyDark,
            topLeft = Offset(w * 0.28f, h * 0.27f),
            size = Size(w * 0.44f, h * 0.20f),
            cornerRadius = CornerRadius(4f, 4f)
        )
        // Прожектор и буферные фонари
        drawCircle(
            color = TrainRTheme.GoldAccent,
            radius = w * 0.05f,
            center = Offset(w * 0.50f, h * 0.58f)
        )
        // Рельсы / шпалы снизу
        drawLine(
            color = TrainRTheme.GoldAccent,
            start = Offset(w * 0.15f, h * 0.82f),
            end = Offset(w * 0.85f, h * 0.82f),
            strokeWidth = 3f
        )
    }
}

@Composable
fun SidebarNavButton(
    title: String,
    icon: androidx.compose.ui.graphics.vector.ImageVector,
    isSelected: Boolean,
    onClick: () -> Unit
) {
    Surface(
        color = if (isSelected) TrainRTheme.ActionBlue else Color.Transparent,
        shape = RoundedCornerShape(8.dp),
        modifier = Modifier
            .fillMaxWidth()
            .clickable { onClick() }
    ) {
        Row(
            verticalAlignment = Alignment.CenterVertically,
            modifier = Modifier.padding(horizontal = 12.dp, vertical = 12.dp)
        ) {
            Icon(
                imageVector = icon,
                contentDescription = title,
                tint = if (isSelected) Color.White else Color.White.copy(alpha = 0.7f),
                modifier = Modifier.size(20.dp)
            )
            Spacer(modifier = Modifier.width(10.dp))
            Text(
                text = title,
                color = if (isSelected) Color.White else Color.White.copy(alpha = 0.75f),
                fontSize = 12.sp,
                fontWeight = if (isSelected) FontWeight.ExtraBold else FontWeight.Medium
            )
        }
    }
}

@Composable
fun TopBarLoadButton(
    label: String,
    count: Int,
    badgeColor: Color,
    onClick: () -> Unit
) {
    OutlinedButton(
        onClick = onClick,
        shape = RoundedCornerShape(8.dp),
        border = BorderStroke(1.dp, TrainRTheme.BorderCol),
        contentPadding = PaddingValues(horizontal = 12.dp, vertical = 8.dp)
    ) {
        Icon(
            imageVector = Icons.Default.UploadFile,
            contentDescription = label,
            tint = TrainRTheme.TextDark,
            modifier = Modifier.size(16.dp)
        )
        Spacer(modifier = Modifier.width(6.dp))
        Text(
            text = label,
            color = TrainRTheme.TextDark,
            fontWeight = FontWeight.Bold,
            fontSize = 12.sp
        )
        Spacer(modifier = Modifier.width(6.dp))
        Surface(
            color = badgeColor,
            shape = RoundedCornerShape(10.dp)
        ) {
            Text(
                text = count.toString(),
                color = Color.White,
                fontSize = 11.sp,
                fontWeight = FontWeight.ExtraBold,
                modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp)
            )
        }
    }
}

@Composable
fun FilterChipItem(
    label: String,
    isSelected: Boolean,
    onClick: () -> Unit
) {
    Surface(
        color = if (isSelected) TrainRTheme.NavyDark else Color.White,
        shape = RoundedCornerShape(16.dp),
        border = BorderStroke(1.dp, if (isSelected) TrainRTheme.NavyDark else TrainRTheme.BorderCol),
        modifier = Modifier.clickable { onClick() }
    ) {
        Text(
            text = label,
            color = if (isSelected) Color.White else TrainRTheme.TextDark,
            fontSize = 12.sp,
            fontWeight = FontWeight.SemiBold,
            modifier = Modifier.padding(horizontal = 14.dp, vertical = 6.dp)
        )
    }
}

@Composable
fun TrainTableSection(
    activeStation: String,
    directionTab: DirectionTab,
    onDirectionChange: (DirectionTab) -> Unit,
    trainFilter: TrainFilterType,
    onTrainFilterChange: (TrainFilterType) -> Unit,
    rows: List<TrainChangeRow>,
    onRowClick: (TrainChangeRow) -> Unit
) {
    Surface(
        color = Color.White,
        shape = RoundedCornerShape(10.dp),
        border = BorderStroke(1.dp, TrainRTheme.BorderCol),
        modifier = Modifier.fillMaxSize()
    ) {
        Column(modifier = Modifier.fillMaxSize()) {
            // Шапка таблицы с переключателем направлений и чипами фильтрации
            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(14.dp),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    DirectionTab.entries.forEach { tab ->
                        val selected = tab == directionTab
                        Button(
                            onClick = { onDirectionChange(tab) },
                            colors = ButtonDefaults.buttonColors(
                                containerColor = if (selected) TrainRTheme.RoyalBlue else Color(0xFFF1F5F9),
                                contentColor = if (selected) Color.White else TrainRTheme.TextDark
                            ),
                            shape = RoundedCornerShape(8.dp)
                        ) {
                            Text(text = tab.title, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }
                    }
                }

                Row(
                    horizontalArrangement = Arrangement.spacedBy(8.dp),
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Text(
                        text = "Станция: $activeStation • Составы:",
                        fontSize = 12.sp,
                        color = TrainRTheme.TextSub,
                        fontWeight = FontWeight.Medium
                    )
                    TrainFilterType.entries.forEach { filter ->
                        FilterChipItem(
                            label = filter.label,
                            isSelected = trainFilter == filter,
                            onClick = { onTrainFilterChange(filter) }
                        )
                    }
                }
            }

            if (activeStation == "Левобережная" && !directionTab.isToMoscow) {
                Box(
                    modifier = Modifier
                        .fillMaxWidth()
                        .background(TrainRTheme.AmberWarn.copy(alpha = 0.15f))
                        .padding(12.dp)
                ) {
                    Text(
                        text = "Инфраструктурное ограничение: На ст. Левобережная платформа №2 (путь №2 «от Москвы») закрыта. Поезда в сторону области проследуют без остановки, изменения не выводятся.",
                        color = TrainRTheme.TextDark,
                        fontWeight = FontWeight.Bold,
                        fontSize = 12.sp
                    )
                }
            }

            // Заголовок колонок таблицы
            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .background(TrainRTheme.NavyDark)
                    .padding(horizontal = 12.dp, vertical = 10.dp),
                verticalAlignment = Alignment.CenterVertically
            ) {
                Text("ДАТЫ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.9f))
                Text("№ ПОЕЗДА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
                Text("МАРШРУТ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.3f))
                Text("ТЕЛЕГРАММА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.4f))
                Text("АФИША", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.7f))
                Text("ТЕЛЕГР.", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.7f))
                Text("ФАКТИЧЕСКИЕ ИЗМЕНЕНИЯ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(2.6f))
                Text("СВЕРКА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
            }

            if (rows.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    Text(
                        text = "Нет данных об изменениях по станции «$activeStation» для выбранного направления и фильтра.",
                        color = TrainRTheme.TextSub,
                        fontSize = 14.sp
                    )
                }
            } else {
                LazyColumn(modifier = Modifier.fillMaxSize()) {
                    items(rows, key = { it.id }) { row ->
                        val bgRowColor = when {
                            row.globalErrorNote != null -> TrainRTheme.RedCancel.copy(alpha = 0.14f)
                            row.isCanceled -> TrainRTheme.RedCancel.copy(alpha = 0.06f)
                            else -> Color.White
                        }
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .background(bgRowColor)
                                .clickable { onRowClick(row) }
                                .padding(horizontal = 12.dp, vertical = 10.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text(
                                text = row.dateStr,
                                fontSize = 12.sp,
                                fontWeight = FontWeight.SemiBold,
                                color = TrainRTheme.TextDark,
                                modifier = Modifier.weight(0.9f)
                            )

                            Row(
                                verticalAlignment = Alignment.CenterVertically,
                                modifier = Modifier.weight(1.1f)
                            ) {
                                Text(
                                    text = "№${row.trainNumber}",
                                    fontSize = 13.sp,
                                    fontWeight = FontWeight.ExtraBold,
                                    color = TrainRTheme.RoyalBlue
                                )
                                if (row.isMcd) {
                                    Spacer(modifier = Modifier.width(6.dp))
                                    Surface(
                                        color = TrainRTheme.PurpleMcd,
                                        shape = RoundedCornerShape(4.dp)
                                    ) {
                                        Text(
                                            text = "МЦД",
                                            color = Color.White,
                                            fontSize = 9.sp,
                                            fontWeight = FontWeight.Bold,
                                            modifier = Modifier.padding(horizontal = 4.dp, vertical = 1.dp)
                                        )
                                    }
                                } else if (row.isExpress) {
                                    Spacer(modifier = Modifier.width(6.dp))
                                    Surface(
                                        color = TrainRTheme.GoldAccent,
                                        shape = RoundedCornerShape(4.dp)
                                    ) {
                                        Text(
                                            text = "ЭКСП",
                                            color = Color.White,
                                            fontSize = 9.sp,
                                            fontWeight = FontWeight.Bold,
                                            modifier = Modifier.padding(horizontal = 4.dp, vertical = 1.dp)
                                        )
                                    }
                                }
                            }

                            Text(
                                text = row.route,
                                fontSize = 12.sp,
                                color = TrainRTheme.TextDark,
                                maxLines = 1,
                                overflow = TextOverflow.Ellipsis,
                                modifier = Modifier.weight(1.3f)
                            )

                            Text(
                                text = row.telegramBasis,
                                fontSize = 11.sp,
                                color = TrainRTheme.TextSub,
                                maxLines = 2,
                                overflow = TextOverflow.Ellipsis,
                                modifier = Modifier.weight(1.4f)
                            )

                            Text(
                                text = row.afishaTime,
                                fontSize = 12.sp,
                                fontFamily = FontFamily.Monospace,
                                color = TrainRTheme.TextSub,
                                modifier = Modifier.weight(0.7f)
                            )

                            Text(
                                text = row.telegramTime,
                                fontSize = 12.sp,
                                fontFamily = FontFamily.Monospace,
                                fontWeight = FontWeight.Bold,
                                color = if (row.isCanceled) TrainRTheme.RedCancel else TrainRTheme.GreenSuccess,
                                modifier = Modifier.weight(0.7f)
                            )

                            Column(modifier = Modifier.weight(2.6f).padding(end = 8.dp)) {
                                Text(
                                    text = row.actualChangesText,
                                    fontSize = 12.sp,
                                    color = TrainRTheme.TextDark
                                )
                                if (row.globalErrorNote != null) {
                                    Text(
                                        text = row.globalErrorNote,
                                        fontSize = 11.sp,
                                        fontWeight = FontWeight.Bold,
                                        color = TrainRTheme.RedCancel
                                    )
                                }
                            }

                            Surface(
                                color = row.auditStatus.color.copy(alpha = 0.14f),
                                shape = RoundedCornerShape(6.dp),
                                border = BorderStroke(1.dp, row.auditStatus.color),
                                modifier = Modifier.weight(1.1f)
                            ) {
                                Text(
                                    text = row.auditStatus.label,
                                    color = row.auditStatus.color,
                                    fontSize = 10.sp,
                                    fontWeight = FontWeight.ExtraBold,
                                    textAlign = TextAlign.Center,
                                    modifier = Modifier.padding(horizontal = 6.dp, vertical = 4.dp)
                                )
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
fun StationAnalyticsReportView(
    activeStation: String,
    rows: List<TrainChangeRow>,
    onRowClick: (TrainChangeRow) -> Unit
) {
    val totalCount = rows.size
    val canceledCount = rows.count { it.isCanceled }
    val mismatchCount = rows.count { it.auditStatus != AuditStatus.OK }
    val mcdCount = rows.count { it.isMcd }

    Column(modifier = Modifier.fillMaxSize()) {
        Row(
            modifier = Modifier.fillMaxWidth(),
            horizontalArrangement = Arrangement.spacedBy(12.dp)
        ) {
            StatSummaryBox(
                title = "ВСЕГО ИЗМЕНЕНИЙ ($activeStation)",
                value = totalCount.toString(),
                accentColor = TrainRTheme.RoyalBlue,
                modifier = Modifier.weight(1f)
            )
            StatSummaryBox(
                title = "ПОЛНЫЕ ОТМЕНЫ ПОЕЗДОВ",
                value = canceledCount.toString(),
                accentColor = TrainRTheme.RedCancel,
                modifier = Modifier.weight(1f)
            )
            StatSummaryBox(
                title = "ПОЕЗДА МЦД-3 (ИВОЛГА / ЭП2Д)",
                value = mcdCount.toString(),
                accentColor = TrainRTheme.PurpleMcd,
                modifier = Modifier.weight(1f)
            )
            StatSummaryBox(
                title = "РАСХОЖДЕНИЯ 4 ИСТОЧНИКОВ",
                value = mismatchCount.toString(),
                accentColor = TrainRTheme.AmberWarn,
                modifier = Modifier.weight(1f)
            )
        }

        Spacer(modifier = Modifier.height(12.dp))

        Surface(
            color = Color.White,
            shape = RoundedCornerShape(10.dp),
            border = BorderStroke(1.dp, TrainRTheme.BorderCol),
            modifier = Modifier.fillMaxSize()
        ) {
            Column(modifier = Modifier.fillMaxSize().padding(14.dp)) {
                Text(
                    text = "СКВОЗНАЯ 4-СТОРОННЯЯ СВЕРКА (АФИША ⟷ ТЕЛЕГРАММА ⟷ ВЫГРУЗКА ⟷ ЯНДЕКС) • СТ. $activeStation",
                    fontSize = 14.sp,
                    fontWeight = FontWeight.ExtraBold,
                    color = TrainRTheme.NavyDark
                )
                Spacer(modifier = Modifier.height(10.dp))

                LazyColumn(
                    modifier = Modifier.fillMaxSize(),
                    verticalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    items(rows, key = { it.id }) { row ->
                        DetailedReportTrainCard(row = row, onClick = { onRowClick(row) })
                    }
                }
            }
        }
    }
}

@Composable
fun StatSummaryBox(
    title: String,
    value: String,
    accentColor: Color,
    modifier: Modifier = Modifier
) {
    Surface(
        color = Color.White,
        shape = RoundedCornerShape(10.dp),
        border = BorderStroke(1.dp, TrainRTheme.BorderCol),
        modifier = modifier
    ) {
        Column(modifier = Modifier.padding(16.dp)) {
            Text(
                text = title,
                fontSize = 11.sp,
                fontWeight = FontWeight.Bold,
                color = TrainRTheme.TextSub
            )
            Spacer(modifier = Modifier.height(6.dp))
            Text(
                text = value,
                fontSize = 26.sp,
                fontWeight = FontWeight.ExtraBold,
                color = accentColor
            )
        }
    }
}

@Composable
fun DetailedReportTrainCard(
    row: TrainChangeRow,
    onClick: () -> Unit
) {
    Surface(
        color = Color(0xFFF8FAFC),
        shape = RoundedCornerShape(8.dp),
        border = BorderStroke(1.dp, row.auditStatus.color.copy(alpha = 0.5f)),
        modifier = Modifier.fillMaxWidth().clickable { onClick() }
    ) {
        Column(modifier = Modifier.padding(12.dp)) {
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text(
                        text = "Поезд №${row.trainNumber} (${row.route})",
                        fontWeight = FontWeight.ExtraBold,
                        fontSize = 14.sp,
                        color = TrainRTheme.NavyDark
                    )
                    Spacer(modifier = Modifier.width(10.dp))
                    Text(
                        text = "Дата: ${row.dateStr} • Основание: ${row.telegramBasis}",
                        fontSize = 12.sp,
                        color = TrainRTheme.TextSub
                    )
                }
                Surface(
                    color = row.auditStatus.color,
                    shape = RoundedCornerShape(6.dp)
                ) {
                    Text(
                        text = row.auditStatus.label,
                        color = Color.White,
                        fontSize = 10.sp,
                        fontWeight = FontWeight.Bold,
                        modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp)
                    )
                }
            }

            Spacer(modifier = Modifier.height(6.dp))
            Text(
                text = "Цепочка сверки: Афиша ${row.afishaTime} ➔ Телеграмма ${row.telegramTime} ➔ Выгрузка ${row.exportTime} ➔ Яндекс ${row.yandexTime}",
                fontSize = 12.sp,
                fontFamily = FontFamily.Monospace,
                fontWeight = FontWeight.SemiBold,
                color = TrainRTheme.RoyalBlue
            )
            Spacer(modifier = Modifier.height(4.dp))
            Text(
                text = "Строка для Word: ${row.wordExportLine}",
                fontSize = 12.sp,
                color = TrainRTheme.TextDark
            )
        }
    }
}

@Composable
fun YandexAuditTabView(
    activeStation: String,
    rows: List<TrainChangeRow>,
    yandexList: List<YandexParsedTrain>,
    onRowClick: (TrainChangeRow) -> Unit
) {
    Surface(
        color = Color.White,
        shape = RoundedCornerShape(10.dp),
        border = BorderStroke(1.dp, TrainRTheme.BorderCol),
        modifier = Modifier.fillMaxSize()
    ) {
        Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Column {
                    Text(
                        text = "АУДИТ PDF-РАСПИСАНИЯ ЯНДЕКСА ПРОТИВ ТЕЛЕГРАММ РЖД И БАЗОВОЙ АФИШИ",
                        fontSize = 15.sp,
                        fontWeight = FontWeight.ExtraBold,
                        color = TrainRTheme.NavyDark
                    )
                    Text(
                        text = "Станция контроля: $activeStation • Загружено записей из PDF Яндекса: ${yandexList.size}",
                        fontSize = 12.sp,
                        color = TrainRTheme.TextSub
                    )
                }
                Surface(
                    color = TrainRTheme.YandexRed.copy(alpha = 0.12f),
                    shape = RoundedCornerShape(8.dp),
                    border = BorderStroke(1.dp, TrainRTheme.YandexRed)
                ) {
                    Text(
                        text = "Контроль Яндекс.Расписаний",
                        color = TrainRTheme.YandexRed,
                        fontWeight = FontWeight.Bold,
                        fontSize = 12.sp,
                        modifier = Modifier.padding(horizontal = 12.dp, vertical = 6.dp)
                    )
                }
            }

            Spacer(modifier = Modifier.height(12.dp))
            HorizontalDivider(color = TrainRTheme.BorderCol)
            Spacer(modifier = Modifier.height(12.dp))

            LazyColumn(
                modifier = Modifier.fillMaxSize(),
                verticalArrangement = Arrangement.spacedBy(8.dp)
            ) {
                items(rows, key = { it.id }) { row ->
                    DetailedReportTrainCard(row = row, onClick = { onRowClick(row) })
                }
            }
        }
    }
}

@Composable
fun TrainDossierDialog(
    trainRow: TrainChangeRow,
    onDismiss: () -> Unit
) {
    Dialog(
        onDismissRequest = onDismiss,
        properties = DialogProperties(usePlatformDefaultWidth = false)
    ) {
        Surface(
            color = Color.White,
            shape = RoundedCornerShape(12.dp),
            modifier = Modifier.width(860.dp).heightIn(min = 480.dp, max = 760.dp)
        ) {
            Column(modifier = Modifier.fillMaxSize().padding(20.dp)) {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Column {
                        Text(
                            text = "ПОЛНЫЙ ОТЧЕТ ПО ПОЕЗДУ (ДОСЬЕ) — №${trainRow.trainNumber}",
                            fontSize = 18.sp,
                            fontWeight = FontWeight.ExtraBold,
                            color = TrainRTheme.NavyDark
                        )
                        Text(
                            text = "${trainRow.route} • Дата: ${trainRow.dateStr} • ${trainRow.telegramBasis}",
                            fontSize = 12.sp,
                            color = TrainRTheme.TextSub
                        )
                    }
                    IconButton(onClick = onDismiss) {
                        Icon(Icons.Default.Close, contentDescription = "Закрыть")
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                Spacer(modifier = Modifier.height(12.dp))

                Text(
                    text = "Нитка следования по станциям (Афиша ➔ Телеграмма ➔ Отклонение в минутах):",
                    fontWeight = FontWeight.Bold,
                    fontSize = 13.sp,
                    color = TrainRTheme.NavyDark
                )
                Spacer(modifier = Modifier.height(8.dp))

                val stopsToShow = if (trainRow.fullThread.isNotEmpty()) {
                    trainRow.fullThread
                } else {
                    listOf(
                        StationStopEntry(
                            stationName = trainRow.station,
                            afishaTime = trainRow.afishaTime,
                            telegramTime = trainRow.telegramTime,
                            exportTime = trainRow.exportTime,
                            yandexTime = trainRow.yandexTime
                        )
                    )
                }

                Column(
                    modifier = Modifier
                        .weight(1f)
                        .fillMaxWidth()
                        .verticalScroll(rememberScrollState())
                        .border(1.dp, TrainRTheme.BorderCol, RoundedCornerShape(8.dp))
                ) {
                    Row(
                        modifier = Modifier
                            .fillMaxWidth()
                            .background(TrainRTheme.NavyDark)
                            .padding(10.dp)
                    ) {
                        Text("СТАНЦИЯ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.5f))
                        Text("АФИША", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f))
                        Text("ТЕЛЕГРАММА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.2f))
                        Text("РАЗНИЦА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f))
                        Text("СТАТУС ОСТАНОВКИ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.5f))
                    }

                    for (stop in stopsToShow) {
                        val diffStr = calculateMinutesDiff(stop.afishaTime, stop.telegramTime)
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(horizontal = 10.dp, vertical = 8.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text(stop.stationName, fontWeight = FontWeight.Bold, fontSize = 12.sp, modifier = Modifier.weight(1.5f))
                            Text(stop.afishaTime, fontFamily = FontFamily.Monospace, fontSize = 12.sp, modifier = Modifier.weight(1f))
                            Text(
                                text = stop.telegramTime,
                                fontFamily = FontFamily.Monospace,
                                fontWeight = FontWeight.Bold,
                                color = if (stop.isSkippedByTelegram) TrainRTheme.RedCancel else TrainRTheme.GreenSuccess,
                                fontSize = 12.sp,
                                modifier = Modifier.weight(1.2f)
                            )
                            Text(diffStr, fontFamily = FontFamily.Monospace, fontSize = 12.sp, modifier = Modifier.weight(1f))
                            Text(
                                text = when {
                                    stop.isSkippedByTelegram -> "БЕЗ ОСТАНОВКИ (прослед.)"
                                    stop.isAddedByTelegram -> "НАЗНАЧЕНА ОСТАНОВКА"
                                    stop.isTechStop -> "ТЕХ. СТОЯНКА (без пасс.)"
                                    else -> "Посадка / высадка пасс."
                                },
                                fontSize = 11.sp,
                                color = if (stop.isSkippedByTelegram) TrainRTheme.RedCancel else TrainRTheme.TextDark,
                                modifier = Modifier.weight(1.5f)
                            )
                        }
                        HorizontalDivider(color = TrainRTheme.BorderCol)
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))
                Text(
                    text = "Исходная выдержка из оперативной телеграммы / выгрузки:",
                    fontWeight = FontWeight.Bold,
                    fontSize = 12.sp,
                    color = TrainRTheme.NavyDark
                )
                Spacer(modifier = Modifier.height(4.dp))
                Surface(
                    color = Color(0xFFF1F5F9),
                    shape = RoundedCornerShape(8.dp),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Text(
                        text = trainRow.rawTelegramSnippet.ifBlank { trainRow.actualChangesText },
                        fontSize = 12.sp,
                        fontFamily = FontFamily.Monospace,
                        modifier = Modifier.padding(12.dp)
                    )
                }

                Spacer(modifier = Modifier.height(12.dp))
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.End) {
                    Button(
                        onClick = onDismiss,
                        colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark)
                    ) {
                        Text("Закрыть досье", color = Color.White)
                    }
                }
            }
        }
    }
}

@Composable
fun CalendarGridPickerDialog(
    availableDatesWithChanges: Set<String>,
    selectedDates: Set<String>,
    onToggleDate: (String) -> Unit,
    onClearAll: () -> Unit,
    onDismiss: () -> Unit
) {
    val currentMonth = remember { YearMonth.of(2026, 10) }
    val nextMonth = remember { currentMonth.plusMonths(1) }

    Dialog(
        onDismissRequest = onDismiss,
        properties = DialogProperties(usePlatformDefaultWidth = false)
    ) {
        Surface(
            color = Color.White,
            shape = RoundedCornerShape(12.dp),
            modifier = Modifier.width(720.dp).padding(16.dp)
        ) {
            Column(modifier = Modifier.padding(20.dp)) {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Column {
                        Text(
                            text = "КАЛЕНДАРЬ ОПЕРАТИВНЫХ ИЗМЕНЕНИЙ (НА 2 МЕСЯЦА ВПЕРЁД)",
                            fontSize = 16.sp,
                            fontWeight = FontWeight.ExtraBold,
                            color = TrainRTheme.NavyDark
                        )
                        Text(
                            text = "Кликабельны только дни, в которых обнаружены изменения в расписании",
                            fontSize = 12.sp,
                            color = TrainRTheme.TextSub
                        )
                    }
                    IconButton(onClick = onDismiss) {
                        Icon(Icons.Default.Close, contentDescription = "Закрыть")
                    }
                }

                Spacer(modifier = Modifier.height(16.dp))

                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(20.dp)
                ) {
                    MonthGridCalendar(
                        yearMonth = currentMonth,
                        monthGenitiveName = "октября",
                        monthTitle = "Октябрь 2026",
                        availableDatesWithChanges = availableDatesWithChanges,
                        selectedDates = selectedDates,
                        onToggleDate = onToggleDate,
                        modifier = Modifier.weight(1f)
                    )
                    MonthGridCalendar(
                        yearMonth = nextMonth,
                        monthGenitiveName = "ноября",
                        monthTitle = "Ноябрь 2026",
                        availableDatesWithChanges = availableDatesWithChanges,
                        selectedDates = selectedDates,
                        onToggleDate = onToggleDate,
                        modifier = Modifier.weight(1f)
                    )
                }

                Spacer(modifier = Modifier.height(18.dp))
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    TextButton(onClick = onClearAll) {
                        Text("Сбросить выделение (показать все дни)", color = TrainRTheme.RedCancel)
                    }
                    Button(
                        onClick = onDismiss,
                        colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue)
                    ) {
                        Text("Применить (${selectedDates.size})", color = Color.White, fontWeight = FontWeight.Bold)
                    }
                }
            }
        }
    }
}

@Composable
fun MonthGridCalendar(
    yearMonth: YearMonth,
    monthGenitiveName: String,
    monthTitle: String,
    availableDatesWithChanges: Set<String>,
    selectedDates: Set<String>,
    onToggleDate: (String) -> Unit,
    modifier: Modifier = Modifier
) {
    Surface(
        color = Color(0xFFF8FAFC),
        shape = RoundedCornerShape(8.dp),
        border = BorderStroke(1.dp, TrainRTheme.BorderCol),
        modifier = modifier
    ) {
        Column(modifier = Modifier.padding(12.dp)) {
            Text(
                text = monthTitle,
                fontWeight = FontWeight.ExtraBold,
                fontSize = 14.sp,
                color = TrainRTheme.NavyDark,
                modifier = Modifier.padding(bottom = 8.dp)
            )

            val daysInMonth = yearMonth.lengthOfMonth()
            val rows = (1..daysInMonth).chunked(7)

            for (week in rows) {
                Row(
                    modifier = Modifier.fillMaxWidth().padding(vertical = 3.dp),
                    horizontalArrangement = Arrangement.spacedBy(4.dp)
                ) {
                    for (day in week) {
                        val dateKey = "$day $monthGenitiveName"
                        val isClickable = dateKey in availableDatesWithChanges
                        val isSelected = dateKey in selectedDates

                        Box(
                            contentAlignment = Alignment.Center,
                            modifier = Modifier
                                .size(36.dp)
                                .clip(RoundedCornerShape(6.dp))
                                .background(
                                    when {
                                        isSelected -> TrainRTheme.ActionBlue
                                        isClickable -> TrainRTheme.GoldAccent.copy(alpha = 0.25f)
                                        else -> Color.Transparent
                                    }
                                )
                                .border(
                                    width = 1.dp,
                                    color = when {
                                        isSelected -> TrainRTheme.ActionBlue
                                        isClickable -> TrainRTheme.GoldAccent
                                        else -> TrainRTheme.BorderCol
                                    },
                                    shape = RoundedCornerShape(6.dp)
                                )
                                .let { mod ->
                                    if (isClickable) mod.clickable { onToggleDate(dateKey) } else mod
                                }
                        ) {
                            Text(
                                text = day.toString(),
                                fontSize = 12.sp,
                                fontWeight = if (isClickable || isSelected) FontWeight.ExtraBold else FontWeight.Normal,
                                color = when {
                                    isSelected -> Color.White
                                    isClickable -> TrainRTheme.NavyDark
                                    else -> TrainRTheme.TextSub.copy(alpha = 0.45f)
                                }
                            )
                        }
                    }
                }
            }
        }
    }
}

// ============================================================================
// 10. СИСТЕМНЫЙ ДИАЛОГ ВЫБОРА ФАЙЛОВ WINDOWS
// ============================================================================

fun pickFilesDialog(
    title: String,
    extensions: Array<String>,
    multiSelect: Boolean
): List<File> {
    val chooser = JFileChooser().apply {
        dialogTitle = title
        isMultiSelectionEnabled = multiSelect
        fileFilter = FileNameExtensionFilter(
            "Поддерживаемые файлы (${extensions.joinToString(", ") { "*.$it" }})",
            *extensions
        )
    }
    val res = chooser.showOpenDialog(null)
    if (res == JFileChooser.APPROVE_OPTION) {
        return if (multiSelect) {
            chooser.selectedFiles?.toList() ?: emptyList()
        } else {
            listOfNotNull(chooser.selectedFile)
        }
    }
    return emptyList()
}
