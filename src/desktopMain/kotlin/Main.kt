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
import androidx.compose.ui.window.WindowPlacement
import androidx.compose.ui.window.WindowPosition
import androidx.compose.ui.window.WindowState
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
import java.awt.Toolkit
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

enum class ReportFilterType(val code: String, val title: String) {
    ALL("ALL", "ВСЕГО ИЗМЕНЕНИЙ"),
    FULL_CANCEL("FULL_CANCEL", "ПОЛНЫЕ ОТМЕНЫ ПОЕЗДОВ"),
    PARTIAL_CANCEL("PARTIAL_CANCEL", "ЧАСТИЧНЫЕ / БЕЗ ОСТАНОВКИ"),
    MCD("MCD", "ПОЕЗДА МЦД-3 (ИВОЛГА/ЭП2Д)"),
    MISMATCH("MISMATCH", "РАСХОЖДЕНИЯ ИСТОЧНИКОВ")
}

enum class AuditStatus(val label: String, val color: Color) {
    OK("СОВПАДАЕТ", Color(0xFF16A34A)),
    ONLY_EXPORT("ТОЛЬКО ВЫГРУЗКА (БЕЗ ТГ)", Color(0xFF2563EB)),
    ONLY_TELEGRAM("ТОЛЬКО ТГ (НЕТ В ВЫГРУЗКЕ)", Color(0xFF0284C7)),
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
    val hasTelegram: Boolean = false,
    val telegramNumber: String = "",
    val telegramDate: String = "",
    val telegramPage: String = "",
    val telegramBasis: String = "—",
    val exportBasis: String = "—",
    val afishaTime: String = "—",
    val telegramTime: String = "—",
    val exportTime: String = "—",
    val yandexTime: String = "—",
    val actualChangesText: String = "",
    val wordExportLine: String = "",
    val isFullCancel: Boolean = false,
    val isCanceled: Boolean = false,
    val isPartialCancel: Boolean = false,
    val isSkippedAtStation: Boolean = false,
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

data class PageTextChunk(
    val pageNumber: Int,
    val text: String
)

data class MonthSpec(
    val monthNumber: Int,
    val genitiveName: String,
    val titleName: String
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
        "Чуприяновка", "Лазурная", "Тверь", "Конаково ГРЭС", "Донховка", "Конаковский Мох", "Путепроводная"
    )

    val MCD3_STATIONS: Set<String> = setOf(
        "Останкино", "Лихоборы", "Моссельмаш", "Грачёвская", "Грачевская",
        "Ховрино", "Левобережная", "Химки", "Молжаниново", "Новоподрезково",
        "Подрезково", "Сходня", "Фирсановская", "Крюково", "Зеленоград-Крюково"
    )

    val HARD_CLOSED_STATIONS: Set<String> = setOf(
        "Малино", "Рижская", "Петровско-Разумовская", "Лазурная"
    )

    val ALL_MONTHS: List<MonthSpec> = listOf(
        MonthSpec(1, "января", "Январь"),
        MonthSpec(2, "февраля", "Февраль"),
        MonthSpec(3, "марта", "Март"),
        MonthSpec(4, "апреля", "Апрель"),
        MonthSpec(5, "мая", "Май"),
        MonthSpec(6, "июня", "Июнь"),
        MonthSpec(7, "июля", "Июль"),
        MonthSpec(8, "августа", "Август"),
        MonthSpec(9, "сентября", "Сентябрь"),
        MonthSpec(10, "октября", "Октябрь"),
        MonthSpec(11, "ноября", "Ноябрь"),
        MonthSpec(12, "декабря", "Декабрь")
    )

    fun normalizeStationName(raw: String): String {
        val clean = raw.trim()
            .replace(Regex("^(?:ст\\.|станция|оп|о\\.п\\.|пл\\.|платформа)\\s*", RegexOption.IGNORE_CASE), "")
            .replace("Зеленоград-Крюково", "Крюково", ignoreCase = true)
            .replace("Грачевская", "Грачёвская", ignoreCase = true)
            .replace("Березки-Дачные", "Берёзки-Дачные", ignoreCase = true)
            .replace("Березки Дачные", "Берёзки-Дачные", ignoreCase = true)
            .replace("Поварово 1", "Поварово-1", ignoreCase = true)
            .replace("Поварово-I", "Поварово-1", ignoreCase = true)
            .replace("Москва-Пасс.-Октябрьская", "Москва", ignoreCase = true)
            .replace("Москва-Пасс.", "Москва", ignoreCase = true)
            .replace("Ленинградский вокзал", "Москва", ignoreCase = true)
            .replace("Оп Подрезково", "Подрезково", ignoreCase = true)
            .replace("Бологое-Моск.", "Бологое", ignoreCase = true)
            .replace("Москва-окт", "Москва", ignoreCase = true)
            .replace("Конаково-ГРЭС", "Конаково ГРЭС", ignoreCase = true)
            .trim()
        return ALL_STATIONS.firstOrNull { it.equals(clean, ignoreCase = true) } ?: clean
    }

    fun isTrainToMoscow(trainNumber: String, route: String = ""): Boolean {
        val r = route.lowercase(Locale.getDefault())
        if (r.contains("на москву") || r.contains("до москвы") || r.contains("до ипподрома") || r.contains("на ипподром") || r.contains("до 47 км")) {
            return true
        }
        if (r.contains("из москвы") || r.contains("от москвы") || r.contains("из ипподрома") ||
            r.contains("до твери") || r.contains("до клина") || r.contains("до крюково") ||
            r.contains("до решетниково") || r.contains("до конаково") || r.contains("до подсолнечной") ||
            r.contains("до сходни") || r.contains("до завидово") || r.contains("до бологое")
        ) {
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
// 3. ТОЧКА ВХОДА ПРИЛОЖЕНИЯ (С АДАПТАЦИЕЙ ПОД ЭКРАН И ВИДИМЫМИ КНОПКАМИ ОКНА)
// ============================================================================

fun main() = application {
    val screenSize = remember {
        try {
            Toolkit.getDefaultToolkit().screenSize
        } catch (_: Exception) {
            java.awt.Dimension(1366, 768)
        }
    }
    val safeWidthDp = (screenSize.width * 0.90).toInt().coerceIn(1024, 1440).dp
    val safeHeightDp = (screenSize.height * 0.86).toInt().coerceIn(680, 880).dp

    val windowState = rememberWindowState(
        width = safeWidthDp,
        height = safeHeightDp,
        position = WindowPosition(Alignment.Center)
    )
    Window(
        onCloseRequest = ::exitApplication,
        state = windowState,
        resizable = true,
        title = "Train-R • Аналитический портал сверки расписания и генерации отмен"
    ) {
        MaterialTheme {
            TrainRApp(
                windowState = windowState,
                onExitApplication = ::exitApplication
            )
        }
    }
}

// ============================================================================
// 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ЛОГИКИ И УМНОГО АНАЛИЗА
// ============================================================================

fun getMonthNameGenitive(monthNum: String): String {
    return when (monthNum.trim().padStart(2, '0')) {
        "01" -> "января"
        "02" -> "февраля"
        "03" -> "марта"
        "04" -> "апреля"
        "05" -> "мая"
        "06" -> "июня"
        "07" -> "июля"
        "08" -> "августа"
        "09" -> "сентября"
        "10" -> "октября"
        "11" -> "ноября"
        "12" -> "декабря"
        else -> "октября"
    }
}

fun normalizeMonthGenitive(rawMonth: String): String {
    val m = rawMonth.trim().lowercase(Locale.getDefault())
    return when {
        m.startsWith("янв") -> "января"
        m.startsWith("фев") -> "февраля"
        m.startsWith("мар") -> "марта"
        m.startsWith("апр") -> "апреля"
        m.startsWith("ма") -> "мая"
        m.startsWith("июн") -> "июня"
        m.startsWith("июл") -> "июля"
        m.startsWith("авг") -> "августа"
        m.startsWith("сен") -> "сентября"
        m.startsWith("окт") -> "октября"
        m.startsWith("ноя") -> "ноября"
        m.startsWith("дек") -> "декабря"
        else -> m
    }
}

fun extractSmartDates(text: String): List<String> {
    val result = mutableListOf<String>()
    val monthNamesPattern = "(январ[яе]|феврал[яе]|март[ае]?|апрел[яе]|ма[яе]|июн[яе]|июл[яе]|август[ае]?|сентябр[яе]|октябр[яе]|ноябр[яе]|декабр[яе])"

    // 1. Диапазоны вида "с 19 по 23 октября" или "19-23 октября"
    val rangeWordRegex = Regex("(?i)(?:с\\s+)?(\\d{1,2})\\s*(?:по|[-–—])\\s*(\\d{1,2})\\s+$monthNamesPattern")
    for (m in rangeWordRegex.findAll(text)) {
        val d1 = m.groupValues[1].toIntOrNull() ?: continue
        val d2 = m.groupValues[2].toIntOrNull() ?: continue
        val monthGen = normalizeMonthGenitive(m.groupValues[3])
        if (d1 in 1..31 && d2 in d1..31) {
            for (d in d1..d2) result.add("$d $monthGen")
        }
    }

    // 2. Перечисления вида "19, 20, 21 и 22 октября"
    val listWordRegex = Regex("(?i)(\\d{1,2}(?:\\s*(?:,|и)\\s*\\d{1,2})*)\\s+$monthNamesPattern")
    for (m in listWordRegex.findAll(text)) {
        val daysRaw = m.groupValues[1]
        val monthGen = normalizeMonthGenitive(m.groupValues[2])
        val days = Regex("\\d{1,2}").findAll(daysRaw).mapNotNull { it.value.toIntOrNull() }.filter { it in 1..31 }
        for (d in days) result.add("$d $monthGen")
    }

    // 3. Числовые даты вида "19, 20.10.2026" или "19.10"
    if (result.isEmpty()) {
        val numRegex = Regex("(?<!\\d)(\\d{1,2}(?:\\s*,\\s*\\d{1,2})*)\\.(\\d{2})(?:\\.\\d{2,4})?(?!\\d)")
        for (m in numRegex.findAll(text)) {
            val monthNum = m.groupValues[2].toIntOrNull() ?: continue
            if (monthNum !in 1..12) continue
            val monthGen = getMonthNameGenitive(m.groupValues[2])
            val days = Regex("\\d{1,2}").findAll(m.groupValues[1]).mapNotNull { it.value.toIntOrNull() }.filter { it in 1..31 }
            for (d in days) result.add("$d $monthGen")
        }
    }

    return result.distinct()
}

fun extractSmartRoute(text: String, trainNum: String, isToMoscow: Boolean, afishaTrain: AfishaBaseTrain?): String {
    // Пытаемся найти явный маршрут в тексте: "до ...", "из ... до ...", "Станция1 - Станция2"
    val explicitToRegex = Regex("(?i)(?:из\\s+[А-Яа-яЁё\\-\\s]+?\\s+)?до\\s+([А-Яа-яЁё\\-]+(?:\\s*\\(\\d+\\s*км\\))?|Ипподрома|Москвы|Крюково|Твери|Клина|Подсолнечной|Сходни|Конаково(?:\\s+ГРЭС)?|Решетниково|Завидово|Бологое)")
    val toMatch = explicitToRegex.find(text)
    if (toMatch != null) {
        return toMatch.value.trim().replace(Regex("\\s{2,}"), " ")
    }

    val dashRouteRegex = Regex("(?i)(Москва|Крюково|Тверь|Клин|Конаково(?:\\s+ГРЭС)?|Подсолнечная|Сходня|Решетниково|Ипподром)\\s*[-–—]\\s*(Москва|Крюково|Тверь|Клин|Конаково(?:\\s+ГРЭС)?|Подсолнечная|Сходня|Решетниково|Ипподром)")
    val dashMatch = dashRouteRegex.find(text)
    if (dashMatch != null) {
        return "${dashMatch.groupValues[1].trim()} — ${dashMatch.groupValues[2].trim()}"
    }

    if (text.contains("Ипподром", ignoreCase = true) || text.contains("47 км", ignoreCase = true)) {
        return if (isToMoscow) "до Ипподрома (47 км)" else "до Крюково"
    }
    if (afishaTrain != null && afishaTrain.route.isNotBlank()) {
        return afishaTrain.route
    }
    return if (isToMoscow) "до Москвы" else "из Москвы"
}

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
        val parts = listStr.split(Regex(",|;|\\s+и\\s+", RegexOption.IGNORE_CASE))
        for (p in parts) {
            val cleaned = p.replace(".", "").trim()
            val norm = PolygonConfig.normalizeStationName(cleaned)
            if (norm in PolygonConfig.ALL_STATIONS) {
                result.add(norm)
            } else {
                // Проверяем вхождение названия станции внутри подстроки
                PolygonConfig.ALL_STATIONS.firstOrNull { st ->
                    cleaned.equals(st, ignoreCase = true)
                }?.let { result.add(it) }
            }
        }
    }
    return result.distinct()
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

fun parseDateForSort(dateStr: String): Int {
    val parts = dateStr.trim().split(Regex("\\s+"))
    if (parts.isEmpty()) return 9999
    val day = parts[0].toIntOrNull() ?: 99
    val monthStr = parts.getOrNull(1)?.lowercase(Locale.getDefault()) ?: ""
    val month = when {
        monthStr.startsWith("янв") -> 1
        monthStr.startsWith("фев") -> 2
        monthStr.startsWith("мар") -> 3
        monthStr.startsWith("апр") -> 4
        monthStr.startsWith("ма") -> 5
        monthStr.startsWith("июн") -> 6
        monthStr.startsWith("июл") -> 7
        monthStr.startsWith("авг") -> 8
        monthStr.startsWith("сен") -> 9
        monthStr.startsWith("окт") -> 10
        monthStr.startsWith("ноя") -> 11
        monthStr.startsWith("дек") -> 12
        else -> 99
    }
    return month * 100 + day
}

fun parseTrainTimeForSorting(timeStr: String): Int {
    if (timeStr == "—" || timeStr.isBlank()) return 9999
    val reg = Regex("(\\d{1,2})[:.](\\d{2})")
    val m = reg.find(timeStr) ?: return 9999
    val hours = m.groupValues[1].toInt()
    val minutes = m.groupValues[2].toInt()

    // СУТКИ НА ЖД: Поезда с 00:00 до 03:59 относятся к КОНЦУ предыдущих суток
    val adjustedHours = if (hours in 0..3) hours + 24 else hours
    return adjustedHours * 60 + minutes
}

fun getSortingTime(row: TrainChangeRow): Int {
    val baseRegex = Regex("(?i)отправлением\\s+в\\s+(\\d{1,2}[:.]\\d{2})")
    val baseMatch = baseRegex.find(row.wordExportLine) ?: baseRegex.find(row.actualChangesText)
    if (baseMatch != null) {
        return parseTrainTimeForSorting(baseMatch.groupValues[1])
    }
    val afishaSort = parseTrainTimeForSorting(row.afishaTime)
    if (afishaSort != 9999) return afishaSort

    val expSort = parseTrainTimeForSorting(row.exportTime)
    if (expSort != 9999) return expSort

    return parseTrainTimeForSorting(row.telegramTime)
}

/**
 * Умная компиляция и объединение данных по одному поезду, дате и станции
 * из нескольких телеграмм и нескольких выгрузок без смешивания источников.
 */
fun mergeSameTrainRowsForDay(rows: List<TrainChangeRow>): List<TrainChangeRow> {
    if (rows.isEmpty()) return emptyList()
    val grouped = rows.groupBy { "${it.dateStr}|${it.station}|${it.trainNumber}" }
    val mergedList = mutableListOf<TrainChangeRow>()

    for ((_, sameTrains) in grouped) {
        val base = sameTrains.first()
        val tgRows = sameTrains.filter { it.hasTelegram }
        val expRows = sameTrains.filter { it.rawExportSnippet.isNotBlank() || it.exportBasis != "—" }

        val hasTg = tgRows.isNotEmpty()
        val combinedSkipped = sameTrains.flatMap { it.skippedStops }.distinct()
        val combinedAdded = sameTrains.flatMap { it.addedStops }.distinct()

        val combinedFullCancel = sameTrains.any { it.isFullCancel }
        val combinedPartial = sameTrains.any { it.isPartialCancel }
        val combinedSkippedAtStation = sameTrains.any { it.isSkippedAtStation }
        val combinedCanceled = combinedFullCancel || combinedSkippedAtStation
        val combinedAppointed = sameTrains.any { it.isAppointed }

        val finalSkipped = if (combinedFullCancel) emptyList() else combinedSkipped

        val latestTelegramTime = if (hasTg) {
            tgRows.map { it.telegramTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: "—"
        } else {
            "—"
        }

        val latestExportTime = if (expRows.isNotEmpty()) {
            expRows.map { it.exportTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: "—"
        } else {
            "—"
        }

        val latestYandexTime = sameTrains.map { it.yandexTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: "—"
        val bestAfishaTime = sameTrains.map { it.afishaTime }.firstOrNull { it != "—" && it.isNotBlank() } ?: "—"

        val bestRoute = sameTrains.map { it.route }
            .firstOrNull { it.isNotBlank() && it != "Выгрузка маршрут" && it != "до Москвы" && it != "из Москвы" }
            ?: sameTrains.map { it.route }.firstOrNull { it.isNotBlank() && it != "Выгрузка маршрут" }
            ?: base.route

        val mergedActualText = expRows.map { it.actualChangesText.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: sameTrains.map { it.actualChangesText.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: ""

        val mergedRawTg = tgRows.map { it.rawTelegramSnippet.trim() }
            .filter { it.isNotBlank() }
            .distinct()
            .joinToString("\n\n--- [Доп. телеграмма] ---\n\n")

        val mergedRawExp = expRows.map { it.rawExportSnippet.trim() }
            .filter { it.isNotBlank() }
            .distinct()
            .joinToString("\n")

        val mergedWordLine = expRows.map { it.wordExportLine.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: sameTrains.map { it.wordExportLine.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: ""

        val mergedTgNumber = tgRows.map { it.telegramNumber.trim() }.filter { it.isNotEmpty() }.distinct().joinToString(", ")
        val mergedTgDate = tgRows.map { it.telegramDate.trim() }.filter { it.isNotEmpty() }.distinct().joinToString(", ")
        val mergedTgPage = tgRows.map { it.telegramPage.trim() }.filter { it.isNotEmpty() }.distinct().joinToString(", ")

        val mergedTgBasis = if (hasTg) {
            tgRows.map { it.telegramBasis.trim() }.filter { it.isNotEmpty() && it != "—" }.distinct().joinToString(", ")
                .ifEmpty { "Телеграмма загружена" }
        } else {
            "—"
        }

        val mergedExpBasis = expRows.map { it.exportBasis.trim() }.filter { it.isNotEmpty() && it != "—" }.distinct().joinToString(", ")
            .ifEmpty { "—" }

        val mergedError = sameTrains.mapNotNull { it.globalErrorNote }.distinct().joinToString("; ").ifEmpty { null }

        val evaluatedStatus = evaluateFourWayAuditStatus(
            afishaTime = bestAfishaTime,
            telegramTime = latestTelegramTime,
            exportTime = latestExportTime,
            yandexTime = latestYandexTime,
            hasTelegram = hasTg,
            hasExport = expRows.isNotEmpty(),
            isCanceled = combinedCanceled,
            skippedStops = finalSkipped,
            globalError = mergedError
        )

        mergedList.add(
            base.copy(
                route = bestRoute,
                hasTelegram = hasTg,
                telegramNumber = mergedTgNumber,
                telegramDate = mergedTgDate,
                telegramPage = mergedTgPage,
                telegramBasis = mergedTgBasis,
                exportBasis = mergedExpBasis,
                afishaTime = bestAfishaTime,
                telegramTime = latestTelegramTime,
                exportTime = latestExportTime,
                yandexTime = latestYandexTime,
                actualChangesText = mergedActualText,
                wordExportLine = mergedWordLine,
                isFullCancel = combinedFullCancel,
                isCanceled = combinedCanceled,
                isPartialCancel = combinedPartial,
                isSkippedAtStation = combinedSkippedAtStation,
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
    return mergedList.sortedBy { getSortingTime(it) }
}

fun buildFullThreads(map: StationDateTrainMap, afishaMap: Map<String, AfishaBaseTrain>): StationDateTrainMap {
    val dateTrainRows = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
    for ((_, dateMap) in map) {
        for ((dt, trainMap) in dateMap) {
            for ((tNum, row) in trainMap) {
                dateTrainRows.getOrPut(dt) { mutableMapOf() }.getOrPut(tNum) { mutableListOf() }.add(row)
            }
        }
    }

    val resultMap = mutableMapOf<String, MutableMap<String, MutableMap<String, TrainChangeRow>>>()

    for ((dt, trainMap) in dateTrainRows) {
        for ((tNum, rows) in trainMap) {
            val baseTrain = afishaMap[tNum]
            val isToMsk = rows.firstOrNull()?.isToMoscow ?: true
            val anyHasTelegram = rows.any { it.hasTelegram }
            val anyHasExport = rows.any { it.rawExportSnippet.isNotBlank() || it.exportBasis != "—" }

            val parsedStations = rows.associateBy { it.station }
            val afishaStations = baseTrain?.stationTimes?.keys ?: emptySet()
            val allStations = (parsedStations.keys + afishaStations).distinct()

            val sortedStations = allStations.sortedWith(
                compareBy(
                    { st ->
                        val parsedRow = parsedStations[st]
                        if (parsedRow != null) {
                            getSortingTime(parsedRow)
                        } else {
                            parseTrainTimeForSorting(baseTrain?.stationTimes?.get(st) ?: "99:99")
                        }
                    },
                    { st ->
                        val idx = PolygonConfig.ALL_STATIONS.indexOfFirst { it.equals(st, ignoreCase = true) }
                        if (isToMsk) -idx else idx
                    }
                )
            )

            val threadEntries = sortedStations.map { st ->
                val parsedRow = parsedStations[st]
                if (parsedRow != null) {
                    val isSkipped = parsedRow.isCanceled ||
                        parsedRow.telegramTime.contains("ОТМЕН", true) ||
                        parsedRow.exportTime.contains("ОТМЕН", true)

                    val safeTgTime = if (!parsedRow.hasTelegram) {
                        "—"
                    } else if (parsedRow.telegramTime == rows.first().telegramTime &&
                        parsedRow.station != rows.first().station &&
                        parsedRow.afishaTime != "—"
                    ) {
                        parsedRow.afishaTime
                    } else {
                        parsedRow.telegramTime
                    }

                    val safeExpTime = if (parsedRow.rawExportSnippet.isBlank() && parsedRow.exportBasis == "—") {
                        "—"
                    } else if (parsedRow.exportTime == rows.first().exportTime &&
                        parsedRow.station != rows.first().station &&
                        parsedRow.afishaTime != "—"
                    ) {
                        parsedRow.afishaTime
                    } else {
                        parsedRow.exportTime
                    }

                    StationStopEntry(
                        stationName = st,
                        afishaTime = parsedRow.afishaTime,
                        telegramTime = safeTgTime,
                        exportTime = safeExpTime,
                        yandexTime = parsedRow.yandexTime,
                        isSkippedByTelegram = isSkipped,
                        isAddedByTelegram = parsedRow.isAppointed,
                        isTechStop = false
                    )
                } else {
                    val aTime = baseTrain?.stationTimes?.get(st) ?: "—"
                    StationStopEntry(
                        stationName = st,
                        afishaTime = aTime,
                        telegramTime = if (anyHasTelegram) aTime else "—",
                        exportTime = if (anyHasExport) aTime else "—",
                        yandexTime = "—"
                    )
                }
            }

            for (r in rows) {
                resultMap.getOrPut(r.station) { mutableMapOf() }
                    .getOrPut(dt) { mutableMapOf() }[tNum] = r.copy(fullThread = threadEntries)
            }
        }
    }
    return resultMap
}

fun evaluateFourWayAuditStatus(
    afishaTime: String,
    telegramTime: String,
    exportTime: String,
    yandexTime: String,
    hasTelegram: Boolean,
    hasExport: Boolean,
    isCanceled: Boolean,
    skippedStops: List<String>,
    globalError: String?
): AuditStatus {
    if (!globalError.isNullOrBlank()) return AuditStatus.GLOBAL_INFRA_ERROR
    if (isCanceled && yandexTime != "—" && !yandexTime.contains("ОТМЕН", true)) {
        return AuditStatus.CRITICAL
    }
    if (hasTelegram && hasExport && telegramTime != "—" && exportTime != "—") {
        val tgCanceled = telegramTime.contains("ОТМЕН", true)
        val expCanceled = exportTime.contains("ОТМЕН", true)
        if (tgCanceled != expCanceled) return AuditStatus.CRITICAL
        if (!tgCanceled && telegramTime != exportTime) return AuditStatus.TIME_MISMATCH
    }
    if (hasTelegram && telegramTime != "—" && yandexTime != "—" && telegramTime != yandexTime && !isCanceled) {
        return AuditStatus.TIME_MISMATCH
    }
    if (skippedStops.isNotEmpty() && yandexTime != "—" && yandexTime == afishaTime) {
        return AuditStatus.STOP_ERROR
    }
    if (!hasTelegram && hasExport) return AuditStatus.ONLY_EXPORT
    if (hasTelegram && !hasExport) return AuditStatus.ONLY_TELEGRAM
    return AuditStatus.OK
}

// ============================================================================
// 5. УМНЫЙ ПАРСИНГ ФАЙЛОВ (ТЕЛЕГРАММЫ С НОМЕРАМИ/ДАТАМИ/СТРАНИЦАМИ И ВЫГРУЗКИ)
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
                    val sb = StringBuilder()
                    for (p in doc.paragraphs) {
                        if (p.text.isNotBlank()) sb.appendLine(p.text)
                    }
                    for (table in doc.tables) {
                        for (row in table.rows) {
                            val rowText = row.tableCells.joinToString(" ") { it.text.trim() }.trim()
                            if (rowText.isNotBlank()) sb.appendLine(rowText)
                        }
                    }
                    return sb.toString()
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
        val utf8Str = String(bytes, Charsets.UTF_8).replace("\uFEFF", "")
        val cp1251Str = String(bytes, Charset.forName("windows-1251")).replace("\uFEFF", "")
        val cyrillicCountUtf8 = Regex("[А-Яа-яЁё]").findAll(utf8Str).count()
        val cyrillicCount1251 = Regex("[А-Яа-яЁё]").findAll(cp1251Str).count()

        return if (utf8Str.contains("\uFFFD") || cyrillicCount1251 > cyrillicCountUtf8) cp1251Str else utf8Str
    }

    /**
     * Постраничное чтение документа, чтобы точно знать страницу (Стр. X) для каждого поезда в телеграмме.
     */
    fun readPagedTextChunks(file: File): List<PageTextChunk> {
        val ext = file.extension.lowercase(Locale.getDefault())
        if (ext == "pdf") {
            val pages = mutableListOf<PageTextChunk>()
            Loader.loadPDF(file).use { pdDoc ->
                val total = pdDoc.numberOfPages
                val stripper = PDFTextStripper()
                for (p in 1..total) {
                    stripper.startPage = p
                    stripper.endPage = p
                    val pageText = stripper.getText(pdDoc)
                    if (pageText.isNotBlank()) {
                        pages.add(PageTextChunk(pageNumber = p, text = pageText))
                    }
                }
            }
            if (pages.isNotEmpty()) return pages
        }

        val fullText = readSmartTextFile(file)
        if (fullText.contains("\u000C")) {
            return fullText.split("\u000C").mapIndexedNotNull { idx, part ->
                if (part.isNotBlank()) PageTextChunk(idx + 1, part) else null
            }
        }

        // Проверяем явные маркеры страниц вида "Стр. 5", "Страница 5", "Лист 5"
        val chunks = mutableListOf<PageTextChunk>()
        var currentPage = 1
        val sb = StringBuilder()
        val pageMarkerRegex = Regex("(?i)^\\s*(?:стр\\.?|страница|лист)\\s*(\\d{1,3})\\b")
        for (line in fullText.lines()) {
            val m = pageMarkerRegex.find(line.trim())
            if (m != null) {
                if (sb.isNotBlank()) {
                    chunks.add(PageTextChunk(currentPage, sb.toString()))
                    sb.clear()
                }
                currentPage = m.groupValues[1].toIntOrNull() ?: (currentPage + 1)
            }
            sb.appendLine(line)
        }
        if (sb.isNotBlank()) {
            chunks.add(PageTextChunk(currentPage, sb.toString()))
        }
        return if (chunks.isEmpty()) listOf(PageTextChunk(1, fullText)) else chunks
    }

    fun extractTelegramMetadata(fullText: String, file: File): Pair<String, String> {
        val headerArea = fullText.lines().take(35).joinToString("\n")

        // Поиск номера телеграммы
        val numRegexes = listOf(
            Regex("(?i)(?:телеграмма|тг)\\s*(?:ЦППК|МТППК|ОКТ|РЖД)?\\s*(?:от\\s*[^№\\n]+)?\\s*(?:№|НР|N)\\s*([\\d/\\-А-Яа-яA-Za-z]+)"),
            Regex("(?i)(?:НР|№)\\s*(\\d{1,5}[/\\-][\\d/\\-А-Яа-яA-Za-z]+)"),
            Regex("(?i)телеграмма\\s*(?:№|НР)?\\s*([\\d/\\-А-Яа-яA-Za-z]{2,})")
        )
        var tgNum = ""
        for (r in numRegexes) {
            val m = r.find(headerArea) ?: r.find(fullText)
            if (m != null) {
                tgNum = "№ ${m.groupValues[1].trim()}"
                break
            }
        }
        if (tgNum.isBlank()) {
            tgNum = file.nameWithoutExtension
        }

        // Поиск даты телеграммы
        val dateRegexes = listOf(
            Regex("(?i)от\\s+(\\d{1,2}\\.\\d{2}\\.\\d{2,4}(?:\\s*г\\.?)?)"),
            Regex("(?i)от\\s+(\\d{1,2}\\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*\\d{0,4}\\s*(?:г\\.?)?)"),
            Regex("(?i)\\b(\\d{2}\\.\\d{2}\\.202\\d)\\b")
        )
        var tgDate = ""
        for (r in dateRegexes) {
            val m = r.find(headerArea)
            if (m != null) {
                tgDate = m.groupValues[1].trim()
                break
            }
        }
        if (tgDate.isBlank()) {
            val sdf = java.text.SimpleDateFormat("dd.MM.yyyy", Locale.getDefault())
            tgDate = "${sdf.format(java.util.Date(file.lastModified()))} (дата файла)"
        }

        return Pair(tgNum, tgDate)
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
                                    val cleanTime = Regex("(\\d{1,2}[:.]\\d{2})").find(timeStr)?.groupValues?.get(1)?.replace('.', ':')?.padStart(5, '0')
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
        if (files.isEmpty()) return emptyMap()

        val rawRowsByStationDate = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        val trainRegex = Regex("(?i)(?:нр\\.|№|поезд\\s*№?)\\s*([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)")
        val timePairRegex = Regex("([А-Яа-яЁё\\-\\s\\d]+?)\\s+(\\d{1,2}[.:]\\d{2}(?::\\d{2})?)(?:\\s*[-–—]\\s*(\\d{1,2}[.:]\\d{2}(?::\\d{2})?))?")

        for (file in files) {
            val fullText = readSmartTextFile(file)
            val pagedChunks = readPagedTextChunks(file)
            val (tgNumber, tgDate) = extractTelegramMetadata(fullText, file)
            val isClarification = fullText.contains("Внести разъяснение", ignoreCase = true)

            data class PagedParagraph(val pageNumber: Int, val text: String)
            val paragraphs = mutableListOf<PagedParagraph>()

            for (chunk in pagedChunks) {
                var currentPara = ""
                for (line in chunk.text.lines()) {
                    val trimmed = line.trim()
                    if (trimmed.isEmpty()) continue

                    val isNewTrainLine = Regex("^(?i)(?:\\d+\\)\\s*)?(?:нр\\.|№|поезд)\\s*[1-7]\\d{2,3}").containsMatchIn(trimmed)
                    val isMajorHeader = Regex("^(?i)(?:\\d+\\.|Отменить|Отменяются|Изменить|Назначить|Назначаются)").containsMatchIn(trimmed) &&
                        !Regex("(?i)(?:нр\\.|№)").containsMatchIn(trimmed)

                    if (isNewTrainLine || isMajorHeader) {
                        if (currentPara.isNotEmpty()) paragraphs.add(PagedParagraph(chunk.pageNumber, currentPara))
                        currentPara = trimmed
                    } else {
                        if (currentPara.isNotEmpty()) currentPara += "\n$trimmed"
                        else currentPara = trimmed
                    }
                }
                if (currentPara.isNotEmpty()) paragraphs.add(PagedParagraph(chunk.pageNumber, currentPara))
            }

            var isGlobalCancelMode = false
            var currentDates = mutableListOf<String>()
            var lastSectionHeader = ""

            for (pagedPara in paragraphs) {
                val rawParaText = pagedPara.text.trim()
                val cleanPara = rawParaText.replace("\n", " ").replace("\r", "")
                    .replace(Regex("\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)", RegexOption.IGNORE_CASE), "")
                    .replace(Regex("\\s{2,}"), " ")
                    .trim()
                if (cleanPara.isEmpty()) continue

                val foundDates = extractSmartDates(cleanPara)
                if (foundDates.isNotEmpty()) {
                    currentDates = foundDates.toMutableList()
                }

                if (Regex("(?:^|[^а-яА-ЯёЁ])(отменить|отменяются)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara) &&
                    !Regex("на\\s+участке|остановк", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)
                ) {
                    isGlobalCancelMode = true
                    lastSectionHeader = cleanPara
                } else if (Regex("(?:^|[^а-яА-ЯёЁ])(изменить|назначить|назначаются|изменяется)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanPara)) {
                    isGlobalCancelMode = false
                    lastSectionHeader = cleanPara
                }

                val trainMatches = trainRegex.findAll(cleanPara).map { it.groupValues[1] }.distinct().toList()
                if (trainMatches.isEmpty()) continue

                val effectiveDates = if (currentDates.isNotEmpty()) currentDates else listOf("Без даты")

                val hasPartialRouteCancel = Regex("(?i)(отмен[а-я]*\\s+на\\s+участке|на\\s+участке\\s+[^.;]+\\s+отмен|следует\\s+до\\s+[А-Яа-яЁё\\-]+\\s+вместо|укорочен)").containsMatchIn(cleanPara)
                val explicitFullCancelWord = Regex("(?i)(?:^|[^а-яА-ЯёЁ])(отменяется|отмен[её]н|отменить|полная\\s+отмена)(?:[^а-яА-ЯёЁ]|$)").containsMatchIn(cleanPara) &&
                    !Regex("(?i)отменен[а-я]*\\s*остановк").containsMatchIn(cleanPara) &&
                    !Regex("(?i)без\\s+остановк").containsMatchIn(cleanPara) &&
                    !hasPartialRouteCancel

                val isFullCancel = (isGlobalCancelMode && !hasPartialRouteCancel && !Regex("(?i)измененным\\s+расписанием|отправится\\s+в|назначен[ыа-я]*\\s+остановк").containsMatchIn(cleanPara)) || explicitFullCancelWord
                val isAppointed = Regex("(?i)назначается|назначить|назначен\\s+поезд").containsMatchIn(cleanPara)

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

                val sampleTrainForDir = trainMatches.firstOrNull() ?: ""
                val isToMskSample = PolygonConfig.isTrainToMoscow(sampleTrainForDir, cleanPara)

                for ((idx, sm) in stopMatches.withIndex()) {
                    val rawSt = sm.groupValues[1].trim()
                    val stName = PolygonConfig.normalizeStationName(rawSt)
                    if (stName !in PolygonConfig.ALL_STATIONS) continue

                    val t1 = sm.groupValues[2].replace('.', ':').take(5).padStart(5, '0')
                    val t2 = sm.groupValues[3].replace('.', ':').take(5).let { if (it.isNotEmpty()) it.padStart(5, '0') else "" }
                    val isTech = cleanPara.contains("тех. стоянка", true) || cleanPara.contains("стоянка техническая", true) || cleanPara.contains("без пассажиров", true)

                    val isClosedSt = PolygonConfig.HARD_CLOSED_STATIONS.contains(stName) ||
                        (stName == "Левобережная" && !isToMskSample)

                    if (isClarification) {
                        if (cleanPara.contains("посадка", true)) addedStops.add(stName)
                        else if (cleanPara.contains("техническая", true)) {
                            if (!isClosedSt) skippedStops.add(stName)
                        }
                    } else {
                        if (t2.isNotEmpty()) {
                            if (isTech) {
                                if (!isClosedSt) skippedStops.add(stName)
                            } else {
                                stationDepTimes[stName] = t2
                            }
                        } else {
                            if (idx == 0 || idx == stopMatches.lastIndex || cleanPara.contains("прибыти", true) || cleanPara.contains("отправлени", true)) {
                                stationDepTimes[stName] = t1
                            } else {
                                if (!isClosedSt) skippedStops.add(stName)
                            }
                        }
                    }
                }

                val mentionedStations = PolygonConfig.ALL_STATIONS.filter { cleanPara.contains(it, true) }
                val targetStations = when {
                    stationDepTimes.isNotEmpty() -> (stationDepTimes.keys + mentionedStations).toSet()
                    mentionedStations.isNotEmpty() -> mentionedStations.toSet()
                    isFullCancel || isAppointed -> setOf("Крюково", "Химки", "Сходня", "Клин", "Подсолнечная", "Тверь")
                    else -> emptySet()
                }

                if (targetStations.isEmpty()) continue

                val validSkippedStops = skippedStops.distinct().filter { it !in PolygonConfig.HARD_CLOSED_STATIONS }
                val validAddedStops = addedStops.distinct().filter { it !in PolygonConfig.HARD_CLOSED_STATIONS }
                val pageLabel = "Стр. ${pagedPara.pageNumber}"
                val fullExcerpt = if (lastSectionHeader.isNotBlank() && !rawParaText.startsWith(lastSectionHeader.take(15))) {
                    "[$lastSectionHeader]\n$rawParaText"
                } else {
                    rawParaText
                }

                for (trainNum in trainMatches) {
                    val afishaTrain = afishaMap[trainNum]
                    val isToMsk = PolygonConfig.isTrainToMoscow(trainNum, cleanPara)
                    val routeStr = extractSmartRoute(cleanPara, trainNum, isToMsk, afishaTrain)

                    for (st in targetStations) {
                        if (st == "Левобережная" && !isToMsk) continue

                        val baseTimeMatch = Regex("(?i)отправлением\\s+в\\s+(\\d{1,2}[:.]\\d{2})").find(cleanPara)
                        val newTimeMatch = Regex("(?i)(?:отправится|отпр\\.?|пойдет|следует)\\s+в\\s+(\\d{1,2}[:.]\\d{2})").find(cleanPara)
                        val timesFound = Regex("(\\d{1,2}[:.]\\d{2})").findAll(cleanPara).map { it.groupValues[1].replace('.', ':').padStart(5, '0') }.toList()

                        val extractedBaseTime = baseTimeMatch?.groupValues?.get(1)?.replace('.', ':')?.padStart(5, '0')
                            ?: timesFound.firstOrNull()
                            ?: "—"
                        val afishaT = afishaTrain?.stationTimes?.get(st) ?: extractedBaseTime

                        val hasGeneralSkip = Regex("(?i)проследует\\s+без\\s+остановки").containsMatchIn(cleanPara) &&
                            validSkippedStops.isEmpty() && mentionedStations.size == 1 && mentionedStations.first() == st
                        val isSkippedHere = st in validSkippedStops || hasGeneralSkip
                        val isCanceledAtThisStation = isFullCancel || isSkippedHere

                        val rawTgT = stationDepTimes[st]
                            ?: newTimeMatch?.groupValues?.get(1)?.replace('.', ':')?.padStart(5, '0')
                            ?: afishaT

                        val tgT = if (isCanceledAtThisStation) {
                            val t = if (afishaT != "—") afishaT else if (rawTgT != "—") rawTgT else ""
                            if (t.isNotEmpty()) "$t ОТМЕНА" else "ОТМЕНА"
                        } else {
                            rawTgT
                        }

                        val isMcd = PolygonConfig.isMcdTrain(trainNum, cleanPara)
                        val isExp = PolygonConfig.isExpressTrain(trainNum, cleanPara)

                        val finalSkippedForNote = validSkippedStops.filter { it != st && isFutureStop(st, it, isToMsk) }
                        val finalAddedForNote = validAddedStops.filter { it != st && isFutureStop(st, it, isToMsk) }

                        val remNoteUI = buildRemainingStopsNote(finalSkippedForNote, finalAddedForNote, false, isCanceledAtThisStation, pathStr, forWord = false)
                        val remNoteWord = buildRemainingStopsNote(finalSkippedForNote, finalAddedForNote, false, isCanceledAtThisStation, pathStr, forWord = true)

                        val changeSummary = when {
                            isFullCancel -> "Поезд №$trainNum ПОЛНОСТЬЮ ОТМЕНЁН (ст. $st)"
                            isSkippedHere -> "Поезд №$trainNum проследует ст. $st БЕЗ ОСТАНОВКИ"
                            hasPartialRouteCancel -> "Поезд №$trainNum: частичная отмена / изменение маршрута ($routeStr)"
                            isAppointed -> "Назначается поезд №$trainNum" + (if (tgT != "—") " отпр. в $tgT" else "") + (if (remNoteUI.isNotEmpty()) " ($remNoteUI)" else "")
                            else -> "Поезд №$trainNum" + (if (tgT != "—") " отпр. в $tgT" else "") + (if (remNoteUI.isNotEmpty()) " ($remNoteUI)" else "")
                        }

                        var wordText = if (isCanceledAtThisStation && !hasPartialRouteCancel) {
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

                        if (isCanceledAtThisStation && !hasPartialRouteCancel) {
                            wordText = wordText.replace(Regex("\\s*\\(?проследует\\s+без\\s+остановки\\)?", RegexOption.IGNORE_CASE), "")
                            if (!Regex("отмен", RegexOption.IGNORE_CASE).containsMatchIn(wordText)) {
                                wordText = "$wordText - ОТМЕНЕН"
                            }
                        }

                        val basisSummary = "ТГ $tgNumber от $tgDate ($pageLabel)"

                        for (dateKey in effectiveDates.distinct()) {
                            val rowObj = TrainChangeRow(
                                id = "${dateKey}_${st}_${trainNum}_${System.nanoTime()}",
                                dateStr = dateKey,
                                station = st,
                                trainNumber = trainNum,
                                isToMoscow = isToMsk,
                                isMcd = isMcd,
                                isExpress = isExp,
                                route = routeStr,
                                hasTelegram = true,
                                telegramNumber = tgNumber,
                                telegramDate = tgDate,
                                telegramPage = pageLabel,
                                telegramBasis = basisSummary,
                                exportBasis = "—",
                                afishaTime = afishaT,
                                telegramTime = tgT,
                                exportTime = "—",
                                yandexTime = "—",
                                actualChangesText = changeSummary.trim(),
                                wordExportLine = wordText,
                                isFullCancel = isFullCancel,
                                isCanceled = isCanceledAtThisStation,
                                isPartialCancel = hasPartialRouteCancel,
                                isSkippedAtStation = isSkippedHere,
                                isAppointed = isAppointed,
                                skippedStops = validSkippedStops,
                                addedStops = validAddedStops,
                                globalErrorNote = closedStationError,
                                rawTelegramSnippet = fullExcerpt,
                                rawExportSnippet = "",
                                auditStatus = evaluateFourWayAuditStatus(
                                    afishaTime = afishaT,
                                    telegramTime = tgT,
                                    exportTime = "—",
                                    yandexTime = "—",
                                    hasTelegram = true,
                                    hasExport = false,
                                    isCanceled = isCanceledAtThisStation,
                                    skippedStops = validSkippedStops,
                                    globalError = closedStationError
                                )
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

    /**
     * Одновременная обработка нескольких файлов выгрузок и слияние с данными телеграмм (если они загружены).
     * Если телеграмма НЕ загружена, поля телеграммы остаются пустыми ("—").
     */
    fun parseSummaryExportFiles(
        files: List<File>,
        afishaMap: Map<String, AfishaBaseTrain>,
        existingTelegramMap: StationDateTrainMap
    ): StationDateTrainMap {
        if (files.isEmpty()) return existingTelegramMap

        val mutableResult = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        for ((st, dtMap) in existingTelegramMap) {
            val mDt = mutableMapOf<String, MutableList<TrainChangeRow>>()
            for ((dt, trMap) in dtMap) mDt[dt] = trMap.values.toMutableList()
            mutableResult[st] = mDt
        }

        val stationHeaderRegex = Regex("^\\s*([А-Яа-яЁё\\-\\s\\d.]+?)\\s*(?:Изменения.*)?$", RegexOption.IGNORE_CASE)
        val trainNumberRegex = Regex("(?:№|нр\\.?)\\s*([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)", RegexOption.IGNORE_CASE)

        for (file in files) {
            val text = readSmartTextFile(file)
            var currentStation = ""
            var currentDates = mutableListOf<String>()

            // Если имя файла содержит название станции (например "Крюково на Москву.docx"), берем как стартовую станцию
            PolygonConfig.ALL_STATIONS.firstOrNull { file.nameWithoutExtension.contains(it, ignoreCase = true) }?.let {
                currentStation = it
            }

            val paragraphs = mutableListOf<String>()
            var currentPara = ""
            for (line in text.lines()) {
                val trimmed = line.trim()
                if (trimmed.isEmpty()) continue
                if (Regex("(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)", RegexOption.IGNORE_CASE).containsMatchIn(trimmed)) continue

                // Проверяем, не является ли строка заголовком станции
                val cleanedForStCheck = trimmed
                    .replace(Regex("(?i)Изменения\\s+в\\s+графике.*"), "")
                    .replace(Regex("(?i)продолжение\\s+на.*"), "")
                    .trim()
                val potentialSt = PolygonConfig.normalizeStationName(cleanedForStCheck)
                if (potentialSt in PolygonConfig.ALL_STATIONS && !trainNumberRegex.containsMatchIn(trimmed)) {
                    if (currentPara.isNotEmpty()) paragraphs.add(currentPara)
                    paragraphs.add(potentialSt)
                    currentPara = ""
                    continue
                }

                // Проверяем, не является ли строка строкой дат без номера поезда
                val datesInLine = extractSmartDates(trimmed)
                if (datesInLine.isNotEmpty() && !trainNumberRegex.containsMatchIn(trimmed)) {
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
                var cleanParagraph = paragraph
                    .replace(Regex("[.\\s]*Изменений\\s*нет\\.?", RegexOption.IGNORE_CASE), "")
                    .replace(Regex("Бологое[- ]*Моск(?:овское)?\\.?", RegexOption.IGNORE_CASE), "")
                    .replace(Regex("\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)", RegexOption.IGNORE_CASE), "")
                    .replace(Regex("\\.{2,}"), ".")
                    .replace(Regex("\\s{2,}"), " ")
                    .trim()

                if (cleanParagraph.isEmpty()) continue

                val stSoloMatch = stationHeaderRegex.find(cleanParagraph)
                if (stSoloMatch != null && !trainNumberRegex.containsMatchIn(cleanParagraph)) {
                    val potentialStation = PolygonConfig.normalizeStationName(stSoloMatch.groupValues[1])
                    if (potentialStation in PolygonConfig.ALL_STATIONS) {
                        currentStation = potentialStation
                        continue
                    }
                }

                val datesFound = extractSmartDates(cleanParagraph)
                if (datesFound.isNotEmpty() && !trainNumberRegex.containsMatchIn(cleanParagraph)) {
                    currentDates = datesFound.toMutableList()
                    continue
                } else if (datesFound.isNotEmpty() && currentDates.isEmpty()) {
                    currentDates = datesFound.toMutableList()
                }

                if (Regex("(?:^|[^а-яА-ЯёЁ])(отменить|отменяются)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph) &&
                    !Regex("на\\s+участке|остановк", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph) &&
                    !trainNumberRegex.containsMatchIn(cleanParagraph)
                ) {
                    isGlobalCancelMode = true
                    continue
                } else if (Regex("(?:^|[^а-яА-ЯёЁ])(изменить|назначить)(?:[^а-яА-ЯёЁ]|$)", RegexOption.IGNORE_CASE).containsMatchIn(cleanParagraph) &&
                    !trainNumberRegex.containsMatchIn(cleanParagraph)
                ) {
                    isGlobalCancelMode = false
                    continue
                }

                val effectiveStation = currentStation.ifEmpty { "Крюково" }
                val effectiveDates = if (currentDates.isNotEmpty()) currentDates else listOf("Без даты")

                val trainMatches = trainNumberRegex.findAll(cleanParagraph).map { it.groupValues[1] }.distinct().toList()
                if (trainMatches.isEmpty()) continue

                val hasPartialRouteCancel = Regex("(?i)(отмен[а-я]*\\s+на\\s+участке|на\\s+участке\\s+[^.;]+\\s+отмен|следует\\s+до\\s+[А-Яа-яЁё\\-]+\\s+вместо|укорочен)").containsMatchIn(cleanParagraph)
                val localFullCancel = Regex("(?i)(?:^|[^а-яА-ЯёЁ])(отменяется|отмен[её]н|отменить|полная\\s+отмена)(?:[^а-яА-ЯёЁ]|$)").containsMatchIn(cleanParagraph) &&
                    !Regex("(?i)отменен[а-я]*\\s+остановк").containsMatchIn(cleanParagraph) &&
                    !Regex("(?i)без\\s+остановк").containsMatchIn(cleanParagraph) &&
                    !hasPartialRouteCancel

                val isFullCancel = (isGlobalCancelMode && !hasPartialRouteCancel) || localFullCancel
                val isAppointed = Regex("(?i)назначается|назначен\\s+поезд").containsMatchIn(cleanParagraph)

                cleanParagraph = cleanParagraph.replace(Regex("\\s*\\(приб\\.?\\s*\\d{1,2}[:.]\\d{2}\\)", RegexOption.IGNORE_CASE), "").trim()

                val skippedStops = mutableListOf<String>()
                skippedStops.addAll(extractStationsByRegex(cleanParagraph, "(?i)отменен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)"))
                skippedStops.addAll(extractStationsByRegex(cleanParagraph, "(?i)без\\s*останов[ок]+\\s*:?\\s*([^.;(]+)"))

                val addedStops = extractStationsByRegex(cleanParagraph, "(?i)назначен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)")

                for (trNum in trainMatches) {
                    val afishaTrain = afishaMap[trNum]
                    val isToMsk = PolygonConfig.isTrainToMoscow(trNum, cleanParagraph)
                    if (effectiveStation == "Левобережная" && !isToMsk) continue

                    val routeStr = extractSmartRoute(cleanParagraph, trNum, isToMsk, afishaTrain)

                    val hasGeneralSkip = Regex("(?i)проследует\\s+без\\s+остановки").containsMatchIn(cleanParagraph) && skippedStops.isEmpty()
                    val isSkippedHere = effectiveStation in skippedStops || hasGeneralSkip
                    val isCanceledAtThisStation = isFullCancel || isSkippedHere

                    val baseTimeMatch = Regex("(?i)отправлением\\s+в\\s+(\\d{1,2}[:.]\\d{2})").find(cleanParagraph)
                    val newTimeMatch = Regex("(?i)(?:отправится|отпр\\.?|пойдет|следует)\\s+в\\s+(\\d{1,2}[:.]\\d{2})").find(cleanParagraph)
                    val timesFound = Regex("(\\d{1,2}[:.]\\d{2})").findAll(cleanParagraph).map { it.groupValues[1].replace('.', ':').padStart(5, '0') }.toList()

                    val extractedBaseTime = baseTimeMatch?.groupValues?.get(1)?.replace('.', ':')?.padStart(5, '0')
                        ?: timesFound.firstOrNull()
                        ?: "—"
                    val afishaT = afishaTrain?.stationTimes?.get(effectiveStation) ?: extractedBaseTime

                    val rawExpT = newTimeMatch?.groupValues?.get(1)?.replace('.', ':')?.padStart(5, '0')
                        ?: if (timesFound.size >= 2) timesFound[1] else if (timesFound.size == 1) timesFound[0] else afishaT

                    val expT = if (isCanceledAtThisStation) {
                        val t = if (afishaT != "—") afishaT else if (rawExpT != "—") rawExpT else ""
                        if (t.isNotEmpty()) "$t ОТМЕНА" else "ОТМЕНА"
                    } else {
                        rawExpT
                    }

                    var wordText = cleanParagraph.replace(Regex("\\s*\\(?проследует\\s+без\\s+остановки\\)?", RegexOption.IGNORE_CASE), "")
                    if (isCanceledAtThisStation && !hasPartialRouteCancel && !Regex("отмен", RegexOption.IGNORE_CASE).containsMatchIn(wordText)) {
                        wordText = "$wordText - ОТМЕНЕН"
                    }
                    wordText = wordText.replace(Regex("(?i)\\s*(?:Малино|Петровско-Разумовская|Рижская|Лазурная)\\s*$"), "").trim()

                    for (currentDate in effectiveDates.distinct()) {
                        val row = TrainChangeRow(
                            id = "${currentDate}_${effectiveStation}_${trNum}_${System.nanoTime()}",
                            dateStr = currentDate,
                            station = effectiveStation,
                            trainNumber = trNum,
                            isToMoscow = isToMsk,
                            isMcd = PolygonConfig.isMcdTrain(trNum, cleanParagraph),
                            isExpress = PolygonConfig.isExpressTrain(trNum, cleanParagraph),
                            route = routeStr,
                            hasTelegram = false,
                            telegramNumber = "",
                            telegramDate = "",
                            telegramPage = "",
                            telegramBasis = "—",
                            exportBasis = file.name,
                            afishaTime = afishaT,
                            telegramTime = "—",
                            exportTime = expT,
                            yandexTime = "—",
                            actualChangesText = cleanParagraph,
                            wordExportLine = wordText,
                            isFullCancel = isFullCancel,
                            isCanceled = isCanceledAtThisStation,
                            isPartialCancel = hasPartialRouteCancel,
                            isSkippedAtStation = isSkippedHere,
                            isAppointed = isAppointed,
                            skippedStops = skippedStops,
                            addedStops = addedStops,
                            rawTelegramSnippet = "",
                            rawExportSnippet = "[${file.name}] $cleanParagraph",
                            auditStatus = AuditStatus.ONLY_EXPORT
                        )
                        mutableResult.getOrPut(effectiveStation) { mutableMapOf() }.getOrPut(currentDate) { mutableListOf() }.add(row)
                    }
                }
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
        val detectedDate = extractSmartDates(text).firstOrNull() ?: "Без даты"
        val lineRegex = Regex("([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?).*?(\\d{1,2}[:.]\\d{2})")
        for (line in text.lines()) {
            val m = lineRegex.find(line) ?: continue
            val trNum = m.groupValues[1]
            val depTime = m.groupValues[2].replace('.', ':').padStart(5, '0')
            parsed.add(
                YandexParsedTrain(
                    trainNumber = trNum,
                    dateStr = detectedDate,
                    station = "Крюково",
                    departureTime = depTime,
                    statusText = if (Regex("отмен", RegexOption.IGNORE_CASE).containsMatchIn(line)) "Отменён" else "По расписанию"
                )
            )
        }
        return parsed
    }
}

// ============================================================================
// 6. ГЕНЕРАТОР ФАЙЛОВ WORD (.DOCX) — ОБЩАЯ ВЫГРУЗКА В ОДНОМ ФАЙЛЕ И ПО СТАНЦИЯМ
// ============================================================================

object WordCancelGenerator {
    private val arrivalExcludeRegex = Regex("(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)", RegexOption.IGNORE_CASE)
    private val nightNoteRegex = Regex("\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)", RegexOption.IGNORE_CASE)

    private fun isTerminalStationForTrain(station: String, row: TrainChangeRow): Boolean {
        val text = row.wordExportLine
        if (text.contains("прибытием", ignoreCase = true) || text.contains("прибывает", ignoreCase = true) || text.contains("прибудет", ignoreCase = true)) return true

        val destPattern = Regex("(?i)№\\s*[\\d/]+\\s+(?:из\\s+[А-Яа-яЁё\\-]+\\s+)?до\\s+([А-Яа-яЁё\\-]+)")
        val match = destPattern.find(text) ?: destPattern.find(row.route)

        if (match != null) {
            val dest = match.groupValues[1].lowercase(Locale.getDefault())
            return when (station) {
                "Москва" -> dest.startsWith("москв")
                "Клин" -> dest.startsWith("клин")
                "Тверь" -> dest.startsWith("твер")
                "Сходня" -> dest.startsWith("сходн")
                "Подсолнечная" -> dest.startsWith("подсолнечн")
                "Решетниково" -> dest.startsWith("решетников")
                "Крюково", "Зеленоград-Крюково" -> dest.startsWith("крюков")
                "Конаково ГРЭС" -> dest.startsWith("конаков")
                "Ипподром" -> dest.startsWith("ипподром")
                else -> dest.startsWith(station.lowercase(Locale.getDefault()).take(5))
            }
        }
        return false
    }

    fun getDirectionSubtitle(isToMoscow: Boolean): String = if (isToMoscow) {
        "на Москву / Ипподром"
    } else "от Москвы / Ипподрома"

    /**
     * Формирует ЕДИНЫЙ общий файл Word (.docx) со сводной выгрузкой по всем выбранным дням
     * и выбранным станциям (объединяя данные из всех загруженных выгрузок и телеграмм).
     */
    fun generateCombinedSingleWordFile(
        selectedStations: Set<String>,
        selectedDates: Set<String>,
        directionTab: DirectionTab,
        allData: StationDateTrainMap,
        includeBothDirections: Boolean = false
    ): File? {
        val desktopDir = File(System.getProperty("user.home"), "Desktop")
        val targetFolder = File(desktopDir, "Отмены по станциям Word")
        if (!targetFolder.exists()) targetFolder.mkdirs()

        val directionsToExport = if (includeBothDirections) {
            listOf(DirectionTab.TO_MOSCOW, DirectionTab.FROM_MOSCOW)
        } else {
            listOf(directionTab)
        }

        val stationsOrdered = PolygonConfig.ALL_STATIONS.filter { it in selectedStations }.ifEmpty {
            allData.keys.toList()
        }

        val suffix = if (includeBothDirections) {
            "оба направления"
        } else if (directionTab.isToMoscow) {
            "на Москву, Ипподром"
        } else {
            "от Москвы, Ипподрома"
        }

        val outFile = File(targetFolder, "Общая выгрузка ($suffix).docx")
        var anySectionWritten = false

        XWPFDocument().use { doc ->
            for (station in stationsOrdered) {
                val stationDateMap = allData[station] ?: continue
                val datesToProcess = if (selectedDates.isEmpty()) stationDateMap.keys else stationDateMap.keys.intersect(selectedDates)
                if (datesToProcess.isEmpty()) continue
                val sortedDatesToProcess = datesToProcess.sortedBy { parseDateForSort(it) }

                for (dir in directionsToExport) {
                    val isToMoscow = dir.isToMoscow
                    if (station == "Левобережная" && !isToMoscow) continue

                    val signatureToDates = linkedMapOf<String, MutableList<String>>()
                    val dateToCleanLines = mutableMapOf<String, List<String>>()

                    for (dateStr in sortedDatesToProcess) {
                        val trainMap = stationDateMap[dateStr] ?: continue
                        val filteredRows = trainMap.values.filter {
                            it.isToMoscow == isToMoscow &&
                                !arrivalExcludeRegex.containsMatchIn(it.wordExportLine) &&
                                !arrivalExcludeRegex.containsMatchIn(it.actualChangesText) &&
                                !isTerminalStationForTrain(station, it)
                        }
                        if (filteredRows.isEmpty()) continue

                        val sortedLines = filteredRows
                            .sortedBy { getSortingTime(it) }
                            .map { nightNoteRegex.replace(it.wordExportLine, "").trim() }
                            .filter { it.isNotBlank() }
                            .distinct()

                        if (sortedLines.isEmpty()) continue
                        dateToCleanLines[dateStr] = sortedLines
                        val normalizedSignature = sortedLines.joinToString("\n")
                        signatureToDates.getOrPut(normalizedSignature) { mutableListOf() }.add(dateStr)
                    }

                    if (signatureToDates.isEmpty()) continue

                    val subtitle = getDirectionSubtitle(isToMoscow)
                    addCenteredBoldHeader(doc, "$station  Изменения в графике движения эл/поездов", pageBreakBefore = anySectionWritten)
                    addCenteredBoldHeader(doc, subtitle, pageBreakBefore = false)
                    anySectionWritten = true
                    var currentPageLineCount = 2

                    for ((_, groupedDates) in signatureToDates) {
                        val formattedDatesHeader = formatMergedDatesHeader(groupedDates)
                        if (currentPageLineCount >= 45) {
                            addContinuationHeader(doc, station, formattedDatesHeader, subtitle)
                            currentPageLineCount = 2
                        }
                        addCenteredBoldHeader(doc, formattedDatesHeader, pageBreakBefore = false)
                        currentPageLineCount++

                        val sampleDate = groupedDates.first()
                        for (rawLine in dateToCleanLines[sampleDate] ?: emptyList()) {
                            if (currentPageLineCount >= 45) {
                                addContinuationHeader(doc, station, formattedDatesHeader, subtitle)
                                currentPageLineCount = 2
                            }
                            val finalTrainLine = rawLine.trim()
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
                }
            }

            if (!anySectionWritten) return null
            FileOutputStream(outFile).use { doc.write(it) }
        }
        return outFile
    }

    /**
     * Формирует как файлы по каждой выбранной станции (собирая все загруженные выгрузки в один файл на станцию по выбранным дням),
     * так и общий сводный файл Word.
     */
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

            val sortedDatesToProcess = datesToProcess.sortedBy { it: String -> parseDateForSort(it) }

            val signatureToDates = linkedMapOf<String, MutableList<String>>()
            val dateToCleanLines = mutableMapOf<String, List<String>>()

            for (dateStr in sortedDatesToProcess) {
                val trainMap = stationDateMap[dateStr] ?: continue
                val filteredRows = trainMap.values
                    .filter {
                        it.isToMoscow == isToMoscow &&
                            !arrivalExcludeRegex.containsMatchIn(it.wordExportLine) &&
                            !arrivalExcludeRegex.containsMatchIn(it.actualChangesText) &&
                            !isTerminalStationForTrain(station, it)
                    }
                if (filteredRows.isEmpty()) continue

                val sortedLines = filteredRows.sortedBy { row ->
                    getSortingTime(row)
                }.map { nightNoteRegex.replace(it.wordExportLine, "").trim() }.filter { it.isNotBlank() }.distinct()

                if (sortedLines.isEmpty()) continue
                dateToCleanLines[dateStr] = sortedLines
                val normalizedSignature = sortedLines.joinToString("\n")
                signatureToDates.getOrPut(normalizedSignature) { mutableListOf() }.add(dateStr)
            }

            if (signatureToDates.isEmpty()) continue
            val fileNameSuffix = if (isToMoscow) "на Москву, Ипподром" else "от Москвы, Ипподрома"
            val outFile = File(targetFolder, "$station $fileNameSuffix.docx")

            XWPFDocument().use { doc ->
                val subtitle = getDirectionSubtitle(isToMoscow)
                addCenteredBoldHeader(doc, "$station  Изменения в графике движения эл/поездов")
                addCenteredBoldHeader(doc, subtitle)
                var currentPageLineCount = 2

                for ((_, groupedDates) in signatureToDates) {
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
                        val finalTrainLine = rawLine.trim()
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

        // Дополнительно формируем единый общий файл Word по выбранным дням
        val combined = generateCombinedSingleWordFile(selectedStations, selectedDates, directionTab, allData, includeBothDirections = false)
        if (combined != null) {
            generatedCount++
        }

        return Pair(generatedCount, targetFolder)
    }

    private fun addCenteredBoldHeader(doc: XWPFDocument, text: String, pageBreakBefore: Boolean = false) {
        val p = doc.createParagraph()
        if (pageBreakBefore) p.isPageBreak = true
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
        val parsed = dates.sortedBy { it: String -> parseDateForSort(it) }.map { d ->
            val parts = d.trim().split(Regex("\\s+"))
            Pair(parts.firstOrNull() ?: d, parts.drop(1).joinToString(" "))
        }
        return parsed.groupBy({ it.second }, { it.first }).entries.joinToString(", ") { (month, days) ->
            val sortedDays = days.mapNotNull { it.toIntOrNull() }.sorted()
            if (sortedDays.isNotEmpty() && month.isNotBlank()) {
                "${sortedDays.joinToString(", ")} $month"
            } else {
                days.joinToString(", ")
            }
        }
    }
}

// ============================================================================
// 7. UI КОМПОНЕНТЫ
// ============================================================================

@OptIn(ExperimentalLayoutApi::class)
@Composable
fun TrainRApp(
    windowState: WindowState,
    onExitApplication: () -> Unit
) {
    val scope = rememberCoroutineScope()

    var currentNavTab: MainNavTab by remember { mutableStateOf(MainNavTab.UPLOADS) }
    var currentDirectionTab: DirectionTab by remember { mutableStateOf(DirectionTab.TO_MOSCOW) }
    var currentTrainFilter: TrainFilterType by remember { mutableStateOf(TrainFilterType.ALL) }
    var currentReportFilter: ReportFilterType by remember { mutableStateOf(ReportFilterType.ALL) }

    var searchQuery: String by remember { mutableStateOf("") }
    var systemStatusText: String by remember { mutableStateOf("Система готова. Загрузите телеграмму или файлы выгрузки.") }

    var afishaCacheMap: Map<String, AfishaBaseTrain> by remember { mutableStateOf(emptyMap()) }
    var stationScheduleMap: StationDateTrainMap by remember { mutableStateOf(emptyMap()) }
    var yandexTrainsList: List<YandexParsedTrain> by remember { mutableStateOf(emptyList()) }

    var loadedTelegramFiles by remember { mutableStateOf(listOf<File>()) }
    var loadedExportFiles by remember { mutableStateOf(listOf<File>()) }

    var showFileManagerDialog by remember { mutableStateOf(false) }

    var activePreviewStation: String by remember { mutableStateOf("Крюково") }
    var checkedStationsForWord: Set<String> by remember { mutableStateOf(setOf("Крюково", "Химки", "Сходня")) }
    var selectedFilterDates: Set<String> by remember { mutableStateOf(emptySet()) }

    var showCalendarDialog: Boolean by remember { mutableStateOf(false) }
    var selectedTrainForDossier: TrainChangeRow? by remember { mutableStateOf(null) }

    fun refreshData() {
        scope.launch(Dispatchers.IO) {
            val tgMap = if (loadedTelegramFiles.isNotEmpty()) {
                FileParsers.parseTelegramFiles(loadedTelegramFiles, afishaCacheMap)
            } else {
                emptyMap()
            }
            val combinedMap = if (loadedExportFiles.isNotEmpty()) {
                FileParsers.parseSummaryExportFiles(loadedExportFiles, afishaCacheMap, tgMap)
            } else {
                tgMap
            }
            val built = buildFullThreads(combinedMap, afishaCacheMap)
            withContext(Dispatchers.Main) {
                stationScheduleMap = built
                val firstStationWithData = PolygonConfig.ALL_STATIONS.firstOrNull { built[it]?.isNotEmpty() == true }
                if (built[activePreviewStation].isNullOrEmpty() && firstStationWithData != null) {
                    activePreviewStation = firstStationWithData
                }
                val tgInfo = if (loadedTelegramFiles.isEmpty()) "ТГ: не загружена" else "ТГ: ${loadedTelegramFiles.size} шт."
                val expInfo = if (loadedExportFiles.isEmpty()) "Выгрузки: 0" else "Выгрузки: ${loadedExportFiles.size} шт."
                systemStatusText = "Данные скомпилированы ($tgInfo | $expInfo)"
            }
        }
    }

    LaunchedEffect(Unit) {
        if (FileParsers.cachedAfishaFile.exists()) {
            withContext(Dispatchers.IO) {
                val parsed = FileParsers.parseAfishaExcel(FileParsers.cachedAfishaFile)
                if (parsed.isNotEmpty()) {
                    withContext(Dispatchers.Main) {
                        afishaCacheMap = parsed
                        systemStatusText = "Загружен кэш базовой Афиши (${parsed.size} поездов). ТГ и выгрузки еще не загружены."
                    }
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
        datesToInclude.sortedBy { parseDateForSort(it) }.flatMap { dt -> dateMap[dt]?.values ?: emptyList() }
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
                    row.actualChangesText.contains(searchQuery.trim(), ignoreCase = true) ||
                    row.telegramNumber.contains(searchQuery.trim(), ignoreCase = true)
            }
    }

    val reportRows: List<TrainChangeRow> = remember(baseRows, currentReportFilter) {
        baseRows.filter { row ->
            when (currentReportFilter) {
                ReportFilterType.ALL -> true
                ReportFilterType.FULL_CANCEL -> row.isFullCancel
                ReportFilterType.PARTIAL_CANCEL -> !row.isFullCancel && (row.isPartialCancel || row.isSkippedAtStation)
                ReportFilterType.MCD -> row.isMcd
                ReportFilterType.MISMATCH -> row.auditStatus != AuditStatus.OK && row.auditStatus != AuditStatus.ONLY_EXPORT
            }
        }
    }

    Row(modifier = Modifier.fillMaxSize().background(TrainRTheme.BgLight)) {
        // ЛЕВАЯ БОКОВАЯ ПАНЕЛЬ
        Column(
            modifier = Modifier
                .width(265.dp)
                .fillMaxHeight()
                .background(TrainRTheme.SidebarNavy)
                .padding(14.dp)
        ) {
            Column(modifier = Modifier.weight(1f).verticalScroll(rememberScrollState())) {
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    modifier = Modifier.padding(bottom = 16.dp)
                ) {
                    VintageTrainLogo(modifier = Modifier.size(40.dp))
                    Spacer(modifier = Modifier.width(10.dp))
                    Column {
                        Text("Train-R", color = Color.White, fontSize = 19.sp, fontWeight = FontWeight.ExtraBold)
                        Text("МТППК • МЦД-3", color = TrainRTheme.GoldAccent, fontSize = 11.sp, fontWeight = FontWeight.SemiBold)
                    }
                }

                HorizontalDivider(color = Color.White.copy(alpha = 0.12f))
                Spacer(modifier = Modifier.height(12.dp))

                SidebarNavButton(title = MainNavTab.UPLOADS.title, icon = Icons.Default.List, isSelected = currentNavTab == MainNavTab.UPLOADS) { currentNavTab = MainNavTab.UPLOADS }
                Spacer(modifier = Modifier.height(6.dp))
                SidebarNavButton(title = MainNavTab.REPORTS.title, icon = Icons.Default.Info, isSelected = currentNavTab == MainNavTab.REPORTS) { currentNavTab = MainNavTab.REPORTS }
                Spacer(modifier = Modifier.height(6.dp))
                SidebarNavButton(title = MainNavTab.YANDEX.title, icon = Icons.Default.CheckCircle, isSelected = currentNavTab == MainNavTab.YANDEX) { currentNavTab = MainNavTab.YANDEX }
            }

            Column(modifier = Modifier.wrapContentHeight(), verticalArrangement = Arrangement.spacedBy(8.dp)) {
                // Кнопка выгрузки ЕДИНОГО общего файла Word по выбранным дням
                Button(
                    onClick = {
                        scope.launch(Dispatchers.IO) {
                            val outFile = WordCancelGenerator.generateCombinedSingleWordFile(
                                selectedStations = checkedStationsForWord,
                                selectedDates = selectedFilterDates,
                                directionTab = currentDirectionTab,
                                allData = stationScheduleMap,
                                includeBothDirections = false
                            )
                            withContext(Dispatchers.Main) {
                                systemStatusText = if (outFile != null) {
                                    "Создан единый файл Word: ${outFile.name} (Папка на Рабочем столе)"
                                } else {
                                    "Нет данных для выгрузки по выбранным станциям и дням."
                                }
                            }
                        }
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.GreenSuccess),
                    shape = RoundedCornerShape(8.dp),
                    modifier = Modifier.fillMaxWidth(),
                    contentPadding = PaddingValues(vertical = 10.dp, horizontal = 8.dp)
                ) {
                    Icon(Icons.Default.Create, contentDescription = null, tint = Color.White, modifier = Modifier.size(16.dp))
                    Spacer(modifier = Modifier.width(6.dp))
                    Text("ОБЩИЙ WORD В 1 ФАЙЛЕ", color = Color.White, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp)
                }

                // Кнопка выгрузки по станциям + общий файл
                Button(
                    onClick = {
                        scope.launch(Dispatchers.IO) {
                            val (count, folder) = WordCancelGenerator.generateStationWordFiles(
                                selectedStations = checkedStationsForWord,
                                selectedDates = selectedFilterDates,
                                directionTab = currentDirectionTab,
                                allData = stationScheduleMap
                            )
                            withContext(Dispatchers.Main) {
                                systemStatusText = "Сформировано файлов Word (.docx): $count шт. Папка: ${folder.absolutePath}"
                            }
                        }
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue),
                    shape = RoundedCornerShape(8.dp),
                    modifier = Modifier.fillMaxWidth(),
                    contentPadding = PaddingValues(vertical = 10.dp, horizontal = 8.dp)
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
                        selectedFilterDates = emptySet()
                        checkedStationsForWord = emptySet()
                        searchQuery = ""
                        scope.launch(Dispatchers.IO) {
                            try {
                                if (FileParsers.cachedAfishaFile.exists()) FileParsers.cachedAfishaFile.delete()
                            } catch (_: Exception) {}
                        }
                        systemStatusText = "Все данные сброшены, кэш очищен."
                    },
                    shape = RoundedCornerShape(8.dp),
                    colors = ButtonDefaults.outlinedButtonColors(contentColor = TrainRTheme.RedCancel),
                    border = BorderStroke(1.dp, TrainRTheme.RedCancel.copy(alpha = 0.5f)),
                    modifier = Modifier.fillMaxWidth(),
                    contentPadding = PaddingValues(vertical = 8.dp)
                ) {
                    Icon(Icons.Default.Delete, contentDescription = null, modifier = Modifier.size(15.dp))
                    Spacer(modifier = Modifier.width(6.dp))
                    Text("СБРОСИТЬ И УДАЛИТЬ", fontWeight = FontWeight.Bold, fontSize = 11.sp)
                }

                Surface(
                    color = TrainRTheme.SidebarActive,
                    shape = RoundedCornerShape(10.dp),
                    border = BorderStroke(1.dp, Color.White.copy(alpha = 0.1f)),
                    modifier = Modifier.fillMaxWidth().padding(top = 4.dp)
                ) {
                    Column(modifier = Modifier.padding(10.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Box(modifier = Modifier.size(8.dp).clip(CircleShape).background(TrainRTheme.GreenSuccess))
                            Spacer(modifier = Modifier.width(8.dp))
                            Text("СТАТУС СИСТЕМЫ", color = Color.White.copy(alpha = 0.7f), fontSize = 10.sp, fontWeight = FontWeight.Bold)
                        }
                        Spacer(modifier = Modifier.height(4.dp))
                        Text(systemStatusText, color = Color.White, fontSize = 11.sp, lineHeight = 14.sp)
                    }
                }
            }
        }

        // ПРАВАЯ ОСНОВНАЯ ОБЛАСТЬ
        Column(modifier = Modifier.weight(1f).fillMaxHeight()) {
            // ВЕРХНЯЯ ПАНЕЛЬ С КНОПКАМИ ЗАГРУЗКИ И ЯВНЫМИ КНОПКАМИ "СВЕРНУТЬ / РАЗВЕРНУТЬ / ЗАКРЫТЬ"
            Surface(color = Color.White, shadowElevation = 2.dp, modifier = Modifier.fillMaxWidth()) {
                Row(
                    modifier = Modifier.fillMaxWidth().padding(horizontal = 16.dp, vertical = 10.dp),
                    verticalAlignment = Alignment.CenterVertically,
                    horizontalArrangement = Arrangement.SpaceBetween
                ) {
                    OutlinedTextField(
                        value = searchQuery,
                        onValueChange = { searchQuery = it },
                        placeholder = { Text("Поиск № поезда, маршрута...", fontSize = 12.sp, color = TrainRTheme.TextSub) },
                        leadingIcon = { Icon(Icons.Default.Search, contentDescription = null, tint = TrainRTheme.TextSub) },
                        trailingIcon = {
                            if (searchQuery.isNotEmpty()) {
                                IconButton(onClick = { searchQuery = "" }) {
                                    Icon(Icons.Default.Clear, contentDescription = null)
                                }
                            }
                        },
                        singleLine = true,
                        shape = RoundedCornerShape(8.dp),
                        modifier = Modifier.width(230.dp).height(48.dp)
                    )

                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp), verticalAlignment = Alignment.CenterVertically) {
                        OutlinedButton(
                            onClick = { showFileManagerDialog = true },
                            shape = RoundedCornerShape(8.dp),
                            border = BorderStroke(1.dp, TrainRTheme.NavyDark),
                            contentPadding = PaddingValues(horizontal = 10.dp, vertical = 8.dp)
                        ) {
                            Icon(Icons.Default.Settings, contentDescription = null, tint = TrainRTheme.NavyDark, modifier = Modifier.size(15.dp))
                            Spacer(modifier = Modifier.width(5.dp))
                            Text("ФАЙЛЫ (${loadedTelegramFiles.size + loadedExportFiles.size})", color = TrainRTheme.NavyDark, fontWeight = FontWeight.Bold, fontSize = 11.sp)
                        }

                        TopBarLoadButton("АФИША", afishaCacheMap.size, TrainRTheme.RoyalBlue) {
                            val f = pickFilesDialog("Афиша .xlsx", arrayOf("xlsx", "xls"), false).firstOrNull()
                            if (f != null) {
                                scope.launch(Dispatchers.IO) {
                                    val parsed = FileParsers.parseAfishaExcel(FileParsers.ensureCacheAfisha(f))
                                    withContext(Dispatchers.Main) {
                                        afishaCacheMap = parsed
                                        systemStatusText = "Афиша закэширована (${parsed.size} поездов)"
                                    }
                                    refreshData()
                                }
                            }
                        }

                        TopBarLoadButton("ТЕЛЕГРАММА", loadedTelegramFiles.size, TrainRTheme.AmberWarn) {
                            val files = pickFilesDialog("Выберите файлы телеграмм (можно несколько)", arrayOf("docx", "pdf", "txt"), true)
                            if (files.isNotEmpty()) {
                                loadedTelegramFiles = (loadedTelegramFiles + files).distinctBy { it.absolutePath }
                                systemStatusText = "Добавлены телеграммы (${loadedTelegramFiles.size} шт.). Анализ..."
                                refreshData()
                            }
                        }

                        TopBarLoadButton("ВЫГРУЗКИ", loadedExportFiles.size, TrainRTheme.GreenSuccess) {
                            val files = pickFilesDialog("Выберите файлы выгрузок (можно несколько одновременно)", arrayOf("txt", "docx", "pdf"), true)
                            if (files.isNotEmpty()) {
                                loadedExportFiles = (loadedExportFiles + files).distinctBy { it.absolutePath }
                                systemStatusText = "Добавлены выгрузки (${loadedExportFiles.size} шт.). Объединение..."
                                refreshData()
                            }
                        }

                        // ЯВНЫЕ КНОПКИ УПРАВЛЕНИЯ ОКНОМ (СВЕРНУТЬ / РАЗВЕРНУТЬ / ЗАКРЫТЬ)
                        HorizontalDivider(modifier = Modifier.height(28.dp).width(1.dp), color = TrainRTheme.BorderCol)

                        OutlinedButton(
                            onClick = { windowState.isMinimized = true },
                            shape = RoundedCornerShape(8.dp),
                            border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                            contentPadding = PaddingValues(horizontal = 10.dp, vertical = 6.dp)
                        ) {
                            Text("— Свернуть", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextDark)
                        }

                        OutlinedButton(
                            onClick = {
                                windowState.placement = if (windowState.placement == WindowPlacement.Maximized) {
                                    WindowPlacement.Floating
                                } else {
                                    WindowPlacement.Maximized
                                }
                            },
                            shape = RoundedCornerShape(8.dp),
                            border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                            contentPadding = PaddingValues(horizontal = 10.dp, vertical = 6.dp)
                        ) {
                            Text(
                                if (windowState.placement == WindowPlacement.Maximized) "❐ В окно" else "☐ Развернуть",
                                fontSize = 11.sp,
                                fontWeight = FontWeight.Bold,
                                color = TrainRTheme.TextDark
                            )
                        }

                        Button(
                            onClick = onExitApplication,
                            colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RedCancel),
                            shape = RoundedCornerShape(8.dp),
                            contentPadding = PaddingValues(horizontal = 10.dp, vertical = 6.dp)
                        ) {
                            Icon(Icons.Default.Close, contentDescription = "Закрыть", tint = Color.White, modifier = Modifier.size(14.dp))
                            Spacer(modifier = Modifier.width(4.dp))
                            Text("Закрыть", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = Color.White)
                        }
                    }
                }
            }

            // ПАНЕЛЬ ФИЛЬТРА ДАТ И ВЫБОРА СТАНЦИЙ ПОЛИГОНА
            Surface(
                color = Color.White,
                border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                modifier = Modifier.fillMaxWidth().padding(horizontal = 16.dp, vertical = 8.dp),
                shape = RoundedCornerShape(10.dp)
            ) {
                Column(modifier = Modifier.padding(12.dp)) {
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.SpaceBetween,
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Row(horizontalArrangement = Arrangement.spacedBy(10.dp), verticalAlignment = Alignment.CenterVertically) {
                            OutlinedButton(onClick = { showCalendarDialog = true }, shape = RoundedCornerShape(8.dp)) {
                                Icon(Icons.Default.DateRange, contentDescription = null, modifier = Modifier.size(17.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text(
                                    if (selectedFilterDates.isEmpty()) "ФИЛЬТР ПО ДАТЕ (ВСЕ 12 МЕСЯЦЕВ)"
                                    else "ВЫБРАНО ДАТ: ${selectedFilterDates.size} (${selectedFilterDates.take(3).joinToString(", ")}${if (selectedFilterDates.size > 3) "..." else ""})",
                                    fontWeight = FontWeight.Bold,
                                    fontSize = 12.sp
                                )
                            }
                            OutlinedButton(
                                onClick = {
                                    systemStatusText = "Выполнена сверка источников."
                                    currentNavTab = MainNavTab.REPORTS
                                },
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(17.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("ОТЧЕТ И СВЕРКА", fontWeight = FontWeight.Bold, fontSize = 12.sp)
                            }
                        }
                        if (selectedFilterDates.isNotEmpty()) {
                            TextButton(onClick = { selectedFilterDates = emptySet() }) {
                                Text("Сбросить фильтр дат (показать все)", fontSize = 12.sp, color = TrainRTheme.RedCancel, fontWeight = FontWeight.Bold)
                            }
                        }
                    }
                    Spacer(modifier = Modifier.height(8.dp))
                    HorizontalDivider(color = TrainRTheme.BorderCol)
                    Spacer(modifier = Modifier.height(6.dp))

                    Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                        Text(
                            "Станции полигона (чекбокс — для выгрузки в Word, клик по названию — просмотр таблицы):",
                            fontSize = 11.sp,
                            color = TrainRTheme.TextSub,
                            fontWeight = FontWeight.SemiBold
                        )
                        Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                            TextButton(
                                onClick = {
                                    val stationsWithData = PolygonConfig.ALL_STATIONS.filter { stationScheduleMap[it]?.isNotEmpty() == true }.toSet()
                                    checkedStationsForWord = if (stationsWithData.isNotEmpty()) stationsWithData else PolygonConfig.ALL_STATIONS.toSet()
                                },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)
                            ) { Text("Станции с изменениями", fontSize = 11.sp) }
                            TextButton(
                                onClick = { checkedStationsForWord = PolygonConfig.ALL_STATIONS.toSet() },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)
                            ) { Text("Выбрать все", fontSize = 11.sp) }
                            TextButton(
                                onClick = { checkedStationsForWord = PolygonConfig.MCD3_STATIONS.intersect(PolygonConfig.ALL_STATIONS.toSet()) },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)
                            ) { Text("Только МЦД-3", fontSize = 11.sp) }
                            TextButton(
                                onClick = { checkedStationsForWord = emptySet() },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp)
                            ) { Text("Снять все", fontSize = 11.sp) }
                        }
                    }

                    Box(modifier = Modifier.fillMaxWidth().heightIn(max = 115.dp).verticalScroll(rememberScrollState())) {
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
                                        isPreview -> TrainRTheme.ActionBlue.copy(alpha = 0.15f)
                                        hasChanges -> TrainRTheme.GoldAccent.copy(alpha = 0.16f)
                                        else -> Color.White
                                    },
                                    border = BorderStroke(
                                        if (isPreview) 1.5.dp else 1.dp,
                                        when {
                                            isPreview -> TrainRTheme.ActionBlue
                                            hasChanges -> TrainRTheme.GoldAccent
                                            else -> TrainRTheme.BorderCol
                                        }
                                    )
                                ) {
                                    Row(
                                        verticalAlignment = Alignment.CenterVertically,
                                        modifier = Modifier.clickable { activePreviewStation = st }.padding(end = 8.dp, top = 2.dp, bottom = 2.dp)
                                    ) {
                                        Checkbox(
                                            checked = isChecked,
                                            onCheckedChange = { checked ->
                                                checkedStationsForWord = if (checked) checkedStationsForWord + st else checkedStationsForWord - st
                                            },
                                            modifier = Modifier.size(24.dp)
                                        )
                                        Text(
                                            st,
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
            }

            // ОСНОВНОЙ КОНТЕНТ ВКЛАДКИ
            Box(modifier = Modifier.weight(1f).fillMaxWidth().padding(horizontal = 16.dp, vertical = 4.dp)) {
                when (currentNavTab) {
                    MainNavTab.UPLOADS -> TrainTableSection(
                        activeStation = activePreviewStation,
                        directionTab = currentDirectionTab,
                        onDirectionChange = { currentDirectionTab = it },
                        trainFilter = currentTrainFilter,
                        onTrainFilterChange = { currentTrainFilter = it },
                        rows = baseRows,
                        hasLoadedTelegrams = loadedTelegramFiles.isNotEmpty(),
                        onRowClick = { selectedTrainForDossier = it }
                    )
                    MainNavTab.REPORTS -> StationAnalyticsReportView(
                        activeStation = activePreviewStation,
                        directionTab = currentDirectionTab,
                        onDirectionChange = { currentDirectionTab = it },
                        baseRows = baseRows,
                        reportRows = reportRows,
                        currentFilter = currentReportFilter,
                        onFilterChange = { currentReportFilter = it },
                        onRowClick = { selectedTrainForDossier = it }
                    )
                    MainNavTab.YANDEX -> YandexAuditTabView(
                        activeStation = activePreviewStation,
                        rows = baseRows,
                        yandexList = yandexTrainsList,
                        onRowClick = { selectedTrainForDossier = it }
                    )
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

    if (showCalendarDialog) {
        CalendarGridPickerDialog(
            availableDatesWithChanges = allDatesWithChanges,
            selectedDates = selectedFilterDates,
            onToggleDate = { dt ->
                selectedFilterDates = if (dt in selectedFilterDates) selectedFilterDates - dt else selectedFilterDates + dt
            },
            onSelectMultipleDates = { dts ->
                selectedFilterDates = selectedFilterDates + dts
            },
            onClearAll = { selectedFilterDates = emptySet() },
            onDismiss = { showCalendarDialog = false }
        )
    }

    selectedTrainForDossier?.let {
        TrainDossierDialog(it) { selectedTrainForDossier = null }
    }
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
        Surface(
            color = Color.White,
            shape = RoundedCornerShape(12.dp),
            modifier = Modifier.fillMaxWidth(0.55f).heightIn(min = 320.dp, max = 620.dp)
        ) {
            Column(modifier = Modifier.fillMaxWidth().padding(20.dp)) {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Text("МЕНЕДЖЕР ЗАГРУЖЕННЫХ ФАЙЛОВ", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        OutlinedButton(onClick = onDismiss, contentPadding = PaddingValues(horizontal = 10.dp, vertical = 4.dp)) {
                            Text("Свернуть", fontSize = 12.sp)
                        }
                        IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, contentDescription = "Закрыть") }
                    }
                }
                Spacer(modifier = Modifier.height(10.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)

                Column(modifier = Modifier.weight(1f, fill = false).verticalScroll(rememberScrollState()).padding(vertical = 12.dp)) {
                    Text("Телеграммы (${telegrams.size})", fontWeight = FontWeight.Bold, color = TrainRTheme.AmberWarn, fontSize = 13.sp)
                    if (telegrams.isEmpty()) Text("Телеграммы еще не загружены", fontSize = 12.sp, color = TrainRTheme.TextSub)
                    telegrams.forEach { file ->
                        Row(
                            modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp),
                            horizontalArrangement = Arrangement.SpaceBetween,
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text(file.name, fontSize = 12.sp, maxLines = 1, overflow = TextOverflow.Ellipsis, modifier = Modifier.weight(1f))
                            IconButton(onClick = { onRemoveTelegram(file) }, modifier = Modifier.size(24.dp)) {
                                Icon(Icons.Default.Delete, contentDescription = "Удалить", tint = TrainRTheme.RedCancel, modifier = Modifier.size(18.dp))
                            }
                        }
                    }

                    Spacer(modifier = Modifier.height(16.dp))

                    Text("Файлы выгрузок (${exports.size})", fontWeight = FontWeight.Bold, color = TrainRTheme.GreenSuccess, fontSize = 13.sp)
                    if (exports.isEmpty()) Text("Выгрузки еще не загружены", fontSize = 12.sp, color = TrainRTheme.TextSub)
                    exports.forEach { file ->
                        Row(
                            modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp),
                            horizontalArrangement = Arrangement.SpaceBetween,
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text(file.name, fontSize = 12.sp, maxLines = 1, overflow = TextOverflow.Ellipsis, modifier = Modifier.weight(1f))
                            IconButton(onClick = { onRemoveExport(file) }, modifier = Modifier.size(24.dp)) {
                                Icon(Icons.Default.Delete, contentDescription = "Удалить", tint = TrainRTheme.RedCancel, modifier = Modifier.size(18.dp))
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))
                Button(
                    onClick = onDismiss,
                    modifier = Modifier.fillMaxWidth(),
                    colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark)
                ) {
                    Text("ЗАКРЫТЬ ОКНО", color = Color.White, fontWeight = FontWeight.Bold)
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
        val bodyPath = Path().apply {
            moveTo(w * 0.22f, h * 0.20f)
            lineTo(w * 0.78f, h * 0.20f)
            lineTo(w * 0.85f, h * 0.72f)
            lineTo(w * 0.15f, h * 0.72f)
            close()
        }
        drawPath(path = bodyPath, color = Color.White)
        drawRoundRect(color = TrainRTheme.NavyDark, topLeft = Offset(w * 0.28f, h * 0.27f), size = Size(w * 0.44f, h * 0.20f), cornerRadius = CornerRadius(4f, 4f))
        drawCircle(color = TrainRTheme.GoldAccent, radius = w * 0.05f, center = Offset(w * 0.50f, h * 0.58f))
        drawLine(color = TrainRTheme.GoldAccent, start = Offset(w * 0.15f, h * 0.82f), end = Offset(w * 0.85f, h * 0.82f), strokeWidth = 3f)
    }
}

@Composable
fun SidebarNavButton(title: String, icon: androidx.compose.ui.graphics.vector.ImageVector, isSelected: Boolean, onClick: () -> Unit) {
    Surface(
        color = if (isSelected) TrainRTheme.ActionBlue else Color.Transparent,
        shape = RoundedCornerShape(8.dp),
        modifier = Modifier.fillMaxWidth().clickable { onClick() }
    ) {
        Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.padding(horizontal = 12.dp, vertical = 10.dp)) {
            Icon(imageVector = icon, contentDescription = title, tint = if (isSelected) Color.White else Color.White.copy(alpha = 0.7f), modifier = Modifier.size(19.dp))
            Spacer(modifier = Modifier.width(10.dp))
            Text(text = title, color = if (isSelected) Color.White else Color.White.copy(alpha = 0.75f), fontSize = 12.sp, fontWeight = if (isSelected) FontWeight.ExtraBold else FontWeight.Medium)
        }
    }
}

@Composable
fun TopBarLoadButton(label: String, count: Int, badgeColor: Color, onClick: () -> Unit) {
    OutlinedButton(
        onClick = onClick,
        shape = RoundedCornerShape(8.dp),
        border = BorderStroke(1.dp, TrainRTheme.BorderCol),
        contentPadding = PaddingValues(horizontal = 10.dp, vertical = 6.dp)
    ) {
        Icon(imageVector = Icons.Default.Add, contentDescription = label, tint = TrainRTheme.TextDark, modifier = Modifier.size(15.dp))
        Spacer(modifier = Modifier.width(4.dp))
        Text(text = label, color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold, fontSize = 11.sp)
        Spacer(modifier = Modifier.width(5.dp))
        Surface(color = badgeColor, shape = RoundedCornerShape(10.dp)) {
            Text(text = count.toString(), color = Color.White, fontSize = 10.sp, fontWeight = FontWeight.ExtraBold, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp))
        }
    }
}

@Composable
fun FilterChipItem(label: String, isSelected: Boolean, onClick: () -> Unit) {
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
    hasLoadedTelegrams: Boolean,
    onRowClick: (TrainChangeRow) -> Unit
) {
    Surface(color = Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = Modifier.fillMaxSize()) {
        Column(modifier = Modifier.fillMaxSize()) {
            Row(
                modifier = Modifier.fillMaxWidth().padding(12.dp),
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
                            Text(tab.title, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }
                    }
                }
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp), verticalAlignment = Alignment.CenterVertically) {
                    Text("Станция: $activeStation • Составы:", fontSize = 12.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.Medium)
                    TrainFilterType.entries.forEach { filter ->
                        FilterChipItem(filter.label, trainFilter == filter) { onTrainFilterChange(filter) }
                    }
                }
            }
            if (!hasLoadedTelegrams && rows.isNotEmpty()) {
                Box(modifier = Modifier.fillMaxWidth().background(Color(0xFFEFF6FF)).padding(horizontal = 12.dp, vertical = 6.dp)) {
                    Text(
                        "ℹ Телеграмма еще не загружена. Отображаются только данные из загруженных файлов выгрузки (колонка «ТЕЛЕГР.» пуста).",
                        color = TrainRTheme.RoyalBlue,
                        fontWeight = FontWeight.SemiBold,
                        fontSize = 11.sp
                    )
                }
            }
            if (activeStation == "Левобережная" && !directionTab.isToMoscow) {
                Box(modifier = Modifier.fillMaxWidth().background(TrainRTheme.AmberWarn.copy(alpha = 0.15f)).padding(12.dp)) {
                    Text("На ст. Левобережная платформа №2 (путь №2 «от Москвы») закрыта. Изменения не выводятся.", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                }
            }
            Row(
                modifier = Modifier.fillMaxWidth().background(TrainRTheme.NavyDark).padding(horizontal = 12.dp, vertical = 10.dp),
                verticalAlignment = Alignment.CenterVertically
            ) {
                Text("ДАТА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.9f))
                Text("№ ПОЕЗДА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
                Text("МАРШРУТ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.3f))
                Text("АФИША", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.7f))
                Text("ТЕЛЕГР.", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.8f))
                Text("ВЫГРУЗКА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.8f))
                Text("ИЗМЕНЕНИЯ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(2.3f))
                Text("СТАТУС", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
            }
            if (rows.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    Text("Нет данных об изменениях по станции «$activeStation». Загрузите выгрузку или телеграмму.", color = TrainRTheme.TextSub, fontSize = 14.sp)
                }
            } else {
                LazyColumn(modifier = Modifier.fillMaxSize()) {
                    items(rows.sortedWith(compareBy({ parseDateForSort(it.dateStr) }, { getSortingTime(it) })), key = { it.id }) { row ->
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .background(
                                    when {
                                        row.globalErrorNote != null -> TrainRTheme.RedCancel.copy(alpha = 0.14f)
                                        row.isFullCancel -> TrainRTheme.RedCancel.copy(alpha = 0.08f)
                                        row.isSkippedAtStation -> TrainRTheme.AmberWarn.copy(alpha = 0.06f)
                                        else -> Color.White
                                    }
                                )
                                .clickable { onRowClick(row) }
                                .padding(horizontal = 12.dp, vertical = 9.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text(row.dateStr, fontSize = 12.sp, fontWeight = FontWeight.SemiBold, color = TrainRTheme.TextDark, modifier = Modifier.weight(0.9f))
                            Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.weight(1.1f)) {
                                Text("№${row.trainNumber}", fontSize = 13.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.RoyalBlue)
                                if (row.isMcd) {
                                    Spacer(Modifier.width(5.dp))
                                    Surface(color = TrainRTheme.PurpleMcd, shape = RoundedCornerShape(4.dp)) {
                                        Text("МЦД", color = Color.White, fontSize = 9.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(horizontal = 4.dp, vertical = 1.dp))
                                    }
                                }
                            }
                            Text(row.route, fontSize = 12.sp, color = TrainRTheme.TextDark, maxLines = 1, overflow = TextOverflow.Ellipsis, modifier = Modifier.weight(1.3f))
                            Text(row.afishaTime, fontSize = 12.sp, fontFamily = FontFamily.Monospace, color = TrainRTheme.TextSub, modifier = Modifier.weight(0.7f))

                            val displayTgTime = if (row.hasTelegram) row.telegramTime else "—"
                            val isTgCanceled = row.hasTelegram && (row.isCanceled || row.telegramTime.contains("ОТМЕН", true))
                            val isExpCanceled = row.exportTime != "—" && (row.isCanceled || row.exportTime.contains("ОТМЕН", true))

                            Text(
                                displayTgTime,
                                fontSize = 12.sp,
                                fontFamily = FontFamily.Monospace,
                                fontWeight = FontWeight.Bold,
                                color = when {
                                    !row.hasTelegram -> TrainRTheme.TextSub
                                    isTgCanceled -> TrainRTheme.RedCancel
                                    else -> TrainRTheme.GreenSuccess
                                },
                                modifier = Modifier.weight(0.8f)
                            )
                            Text(
                                row.exportTime,
                                fontSize = 12.sp,
                                fontFamily = FontFamily.Monospace,
                                fontWeight = FontWeight.Bold,
                                color = if (isExpCanceled) TrainRTheme.RedCancel else TrainRTheme.TextDark,
                                modifier = Modifier.weight(0.8f)
                            )
                            Column(modifier = Modifier.weight(2.3f).padding(end = 8.dp)) {
                                Text(row.actualChangesText, fontSize = 12.sp, color = TrainRTheme.TextDark)
                                if (row.hasTelegram && row.telegramPage.isNotBlank()) {
                                    Text("ТГ: ${row.telegramNumber} (${row.telegramPage})", fontSize = 10.sp, color = TrainRTheme.AmberWarn, fontWeight = FontWeight.Bold)
                                }
                                if (row.globalErrorNote != null) {
                                    Text(row.globalErrorNote, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RedCancel)
                                }
                            }
                            Surface(
                                color = row.auditStatus.color.copy(alpha = 0.14f),
                                shape = RoundedCornerShape(6.dp),
                                border = BorderStroke(1.dp, row.auditStatus.color),
                                modifier = Modifier.weight(1.1f)
                            ) {
                                Text(
                                    row.auditStatus.label,
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
    directionTab: DirectionTab,
    onDirectionChange: (DirectionTab) -> Unit,
    baseRows: List<TrainChangeRow>,
    reportRows: List<TrainChangeRow>,
    currentFilter: ReportFilterType,
    onFilterChange: (ReportFilterType) -> Unit,
    onRowClick: (TrainChangeRow) -> Unit
) {
    val totalCount = baseRows.size
    val fullCancelCount = baseRows.count { it.isFullCancel }
    val partialCancelCount = baseRows.count { !it.isFullCancel && (it.isPartialCancel || it.isSkippedAtStation) }
    val mcdCount = baseRows.count { it.isMcd }
    val mismatchCount = baseRows.count { it.auditStatus != AuditStatus.OK && it.auditStatus != AuditStatus.ONLY_EXPORT }

    Column(modifier = Modifier.fillMaxSize()) {
        // Кнопки фильтрации по категориям (при нажатии внизу отображаются строго поезда выбранной категории)
        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
            StatSummaryBox(
                title = "ВСЕГО ИЗМЕНЕНИЙ ($activeStation)",
                value = totalCount.toString(),
                accentColor = TrainRTheme.RoyalBlue,
                isSelected = currentFilter == ReportFilterType.ALL,
                onClick = { onFilterChange(ReportFilterType.ALL) },
                modifier = Modifier.weight(1f)
            )
            StatSummaryBox(
                title = "ПОЛНЫЕ ОТМЕНЫ ПОЕЗДОВ",
                value = fullCancelCount.toString(),
                accentColor = TrainRTheme.RedCancel,
                isSelected = currentFilter == ReportFilterType.FULL_CANCEL,
                onClick = { onFilterChange(ReportFilterType.FULL_CANCEL) },
                modifier = Modifier.weight(1f)
            )
            StatSummaryBox(
                title = "ЧАСТИЧНЫЕ / БЕЗ ОСТАНОВКИ",
                value = partialCancelCount.toString(),
                accentColor = TrainRTheme.AmberWarn,
                isSelected = currentFilter == ReportFilterType.PARTIAL_CANCEL,
                onClick = { onFilterChange(ReportFilterType.PARTIAL_CANCEL) },
                modifier = Modifier.weight(1f)
            )
            StatSummaryBox(
                title = "ПОЕЗДА МЦД-3",
                value = mcdCount.toString(),
                accentColor = TrainRTheme.PurpleMcd,
                isSelected = currentFilter == ReportFilterType.MCD,
                onClick = { onFilterChange(ReportFilterType.MCD) },
                modifier = Modifier.weight(1f)
            )
            StatSummaryBox(
                title = "РАСХОЖДЕНИЯ И ОШИБКИ",
                value = mismatchCount.toString(),
                accentColor = Color(0xFFB91C1C),
                isSelected = currentFilter == ReportFilterType.MISMATCH,
                onClick = { onFilterChange(ReportFilterType.MISMATCH) },
                modifier = Modifier.weight(1f)
            )
        }
        Spacer(modifier = Modifier.height(10.dp))
        Surface(color = Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = Modifier.fillMaxSize()) {
            Column(modifier = Modifier.fillMaxSize().padding(14.dp)) {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Column {
                        Text(
                            "ОТЧЕТ: ${currentFilter.title} • СТ. $activeStation (Найдено: ${reportRows.size})",
                            fontSize = 14.sp,
                            fontWeight = FontWeight.ExtraBold,
                            color = TrainRTheme.NavyDark
                        )
                        Text(
                            "Нажмите на любую карточку сверху для фильтрации списка или на поезд для открытия полного досье",
                            fontSize = 11.sp,
                            color = TrainRTheme.TextSub
                        )
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        DirectionTab.entries.forEach { tab ->
                            val selected = tab == directionTab
                            Button(
                                onClick = { onDirectionChange(tab) },
                                colors = ButtonDefaults.buttonColors(
                                    containerColor = if (selected) TrainRTheme.RoyalBlue else Color(0xFFF1F5F9),
                                    contentColor = if (selected) Color.White else TrainRTheme.TextDark
                                ),
                                shape = RoundedCornerShape(8.dp),
                                contentPadding = PaddingValues(horizontal = 10.dp, vertical = 6.dp)
                            ) {
                                Text(tab.title, fontWeight = FontWeight.Bold, fontSize = 11.sp)
                            }
                        }
                    }
                }
                Spacer(modifier = Modifier.height(10.dp))
                if (reportRows.isEmpty()) {
                    Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                        Text(
                            "По фильтру «${currentFilter.title}» на станции «$activeStation» поездов не найдено.",
                            color = TrainRTheme.TextSub,
                            fontSize = 14.sp,
                            fontWeight = FontWeight.SemiBold
                        )
                    }
                } else {
                    LazyColumn(modifier = Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(10.dp)) {
                        items(reportRows.sortedWith(compareBy({ parseDateForSort(it.dateStr) }, { getSortingTime(it) })), key = { it.id }) { row ->
                            DetailedReportTrainCard(row) { onRowClick(row) }
                        }
                    }
                }
            }
        }
    }
}

@Composable
fun StatSummaryBox(title: String, value: String, accentColor: Color, isSelected: Boolean, onClick: () -> Unit, modifier: Modifier = Modifier) {
    Surface(
        color = if (isSelected) accentColor.copy(alpha = 0.12f) else Color.White,
        shape = RoundedCornerShape(10.dp),
        border = BorderStroke(if (isSelected) 2.dp else 1.dp, if (isSelected) accentColor else TrainRTheme.BorderCol),
        modifier = modifier.clickable { onClick() }
    ) {
        Column(modifier = Modifier.padding(12.dp)) {
            Text(title, fontSize = 10.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub, maxLines = 1, overflow = TextOverflow.Ellipsis)
            Spacer(modifier = Modifier.height(4.dp))
            Text(value, fontSize = 24.sp, fontWeight = FontWeight.ExtraBold, color = accentColor)
        }
    }
}

@Composable
fun DetailedReportTrainCard(row: TrainChangeRow, onClick: () -> Unit) {
    Surface(
        color = Color(0xFFF8FAFC),
        shape = RoundedCornerShape(8.dp),
        border = BorderStroke(1.dp, row.auditStatus.color.copy(alpha = 0.5f)),
        modifier = Modifier.fillMaxWidth().clickable { onClick() }
    ) {
        Column(modifier = Modifier.padding(12.dp)) {
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text("Поезд №${row.trainNumber} (${row.route})", fontWeight = FontWeight.ExtraBold, fontSize = 14.sp, color = TrainRTheme.NavyDark)
                    Spacer(modifier = Modifier.width(10.dp))
                    if (row.isFullCancel) {
                        Surface(color = TrainRTheme.RedCancel, shape = RoundedCornerShape(4.dp)) {
                            Text("ПОЛНАЯ ОТМЕНА", color = Color.White, fontSize = 10.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp))
                        }
                        Spacer(modifier = Modifier.width(8.dp))
                    } else if (row.isPartialCancel || row.isSkippedAtStation) {
                        Surface(color = TrainRTheme.AmberWarn, shape = RoundedCornerShape(4.dp)) {
                            Text("БЕЗ ОСТАНОВКИ / ЧАСТИЧНАЯ", color = Color.White, fontSize = 10.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp))
                        }
                        Spacer(modifier = Modifier.width(8.dp))
                    }
                    Text("Дата: ${row.dateStr} • Станция: ${row.station}", fontSize = 12.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.SemiBold)
                }
                Surface(color = row.auditStatus.color, shape = RoundedCornerShape(6.dp)) {
                    Text(row.auditStatus.label, color = Color.White, fontSize = 10.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp))
                }
            }
            Spacer(modifier = Modifier.height(6.dp))
            val tgTimeDisplay = if (row.hasTelegram) row.telegramTime else "— (не загружена)"
            Text(
                "Цепочка сверки: Афиша ${row.afishaTime} ➔ Телеграмма $tgTimeDisplay ➔ Выгрузка ${row.exportTime} ➔ Яндекс ${row.yandexTime}",
                fontSize = 12.sp,
                fontFamily = FontFamily.Monospace,
                fontWeight = FontWeight.SemiBold,
                color = TrainRTheme.RoyalBlue
            )
            Spacer(modifier = Modifier.height(4.dp))
            Text("Строка для Word: ${row.wordExportLine}", fontSize = 12.sp, color = TrainRTheme.TextDark, fontWeight = FontWeight.Medium)

            Spacer(modifier = Modifier.height(8.dp))
            if (row.hasTelegram) {
                Surface(
                    color = Color(0xFFFFFBEB),
                    shape = RoundedCornerShape(6.dp),
                    border = BorderStroke(1.dp, TrainRTheme.GoldAccent.copy(alpha = 0.6f)),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Column(modifier = Modifier.padding(10.dp)) {
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(16.dp)) {
                            Text("Телеграмма: ${row.telegramNumber.ifBlank { "—" }}", fontWeight = FontWeight.ExtraBold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
                            Text("Дата ТГ: ${row.telegramDate.ifBlank { "—" }}", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
                            if (row.telegramPage.isNotBlank()) {
                                Text("Страница: ${row.telegramPage}", fontWeight = FontWeight.ExtraBold, fontSize = 12.sp, color = TrainRTheme.RoyalBlue)
                            }
                        }
                        Spacer(modifier = Modifier.height(4.dp))
                        Text(
                            "Полная выдержка по поезду №${row.trainNumber}:",
                            fontWeight = FontWeight.Bold,
                            fontSize = 11.sp,
                            color = TrainRTheme.TextSub
                        )
                        Spacer(modifier = Modifier.height(2.dp))
                        Text(
                            row.rawTelegramSnippet,
                            fontSize = 11.sp,
                            color = TrainRTheme.TextDark,
                            fontFamily = FontFamily.Monospace
                        )
                    }
                }
            } else {
                Text(
                    "Телеграмма: не загружена (данные получены из файла выгрузки: ${row.exportBasis})",
                    fontSize = 11.sp,
                    color = TrainRTheme.TextSub,
                    fontWeight = FontWeight.Medium
                )
            }
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
    Surface(color = Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = Modifier.fillMaxSize()) {
        Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                Column {
                    Text("АУДИТ PDF-РАСПИСАНИЯ ЯНДЕКСА ПРОТИВ ТЕЛЕГРАММ РЖД И БАЗОВОЙ АФИШИ", fontSize = 15.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                    Text("Станция контроля: $activeStation • Загружено записей из PDF Яндекса: ${yandexList.size}", fontSize = 12.sp, color = TrainRTheme.TextSub)
                }
                Surface(color = TrainRTheme.YandexRed.copy(alpha = 0.12f), shape = RoundedCornerShape(8.dp), border = BorderStroke(1.dp, TrainRTheme.YandexRed)) {
                    Text("Контроль Яндекс.Расписаний", color = TrainRTheme.YandexRed, fontWeight = FontWeight.Bold, fontSize = 12.sp, modifier = Modifier.padding(horizontal = 12.dp, vertical = 6.dp))
                }
            }
            Spacer(modifier = Modifier.height(12.dp))
            HorizontalDivider(color = TrainRTheme.BorderCol)
            Spacer(modifier = Modifier.height(12.dp))
            LazyColumn(modifier = Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(8.dp)) {
                items(rows, key = { it.id }) { row -> DetailedReportTrainCard(row) { onRowClick(row) } }
            }
        }
    }
}

@Composable
fun TrainDossierDialog(trainRow: TrainChangeRow, onDismiss: () -> Unit) {
    Dialog(onDismissRequest = onDismiss, properties = DialogProperties(usePlatformDefaultWidth = false)) {
        Surface(
            color = Color.White,
            shape = RoundedCornerShape(12.dp),
            modifier = Modifier.fillMaxWidth(0.90f).fillMaxHeight(0.88f)
        ) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                // Шапка досье с всегда видимыми кнопками "Свернуть" и "Закрыть"
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column(modifier = Modifier.weight(1f)) {
                        Text("ПОЛНЫЙ ОТЧЕТ ПО ПОЕЗДУ (ДОСЬЕ) — №${trainRow.trainNumber}", fontSize = 17.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                        val tgHeader = if (trainRow.hasTelegram) {
                            "Телеграмма: ${trainRow.telegramNumber} от ${trainRow.telegramDate} (${trainRow.telegramPage})"
                        } else {
                            "Телеграмма: не загружена"
                        }
                        Text("${trainRow.route} • Дата: ${trainRow.dateStr} • $tgHeader", fontSize = 12.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.SemiBold)
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.BorderCol)) {
                            Text("Свернуть", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RedCancel)) {
                            Text("Закрыть", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }
                    }
                }
                Spacer(modifier = Modifier.height(8.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                Spacer(modifier = Modifier.height(8.dp))
                Text("Нитка следования по станциям (Афиша ➔ Телеграмма ➔ Выгрузка ➔ Разница):", fontWeight = FontWeight.Bold, fontSize = 13.sp, color = TrainRTheme.NavyDark)
                Spacer(modifier = Modifier.height(6.dp))

                val stopsToShow = if (trainRow.fullThread.isNotEmpty()) {
                    trainRow.fullThread
                } else {
                    listOf(
                        StationStopEntry(
                            stationName = trainRow.station,
                            afishaTime = trainRow.afishaTime,
                            telegramTime = if (trainRow.hasTelegram) trainRow.telegramTime else "—",
                            exportTime = trainRow.exportTime,
                            yandexTime = trainRow.yandexTime
                        )
                    )
                }

                Column(modifier = Modifier.fillMaxWidth().weight(0.52f).border(1.dp, TrainRTheme.BorderCol, RoundedCornerShape(8.dp))) {
                    Row(modifier = Modifier.fillMaxWidth().background(TrainRTheme.NavyDark).padding(horizontal = 10.dp, vertical = 6.dp)) {
                        Text("СТАНЦИЯ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.5f))
                        Text("АФИША", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(0.9f))
                        Text("ТЕЛЕГРАММА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.1f))
                        Text("ВЫГРУЗКА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.1f))
                        Text("РАЗНИЦА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(0.9f))
                        Text("СТАТУС ОСТАНОВКИ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.5f))
                    }
                    LazyColumn(modifier = Modifier.fillMaxWidth()) {
                        items(stopsToShow) { stop ->
                            val activeComparisonTime = stop.exportTime.takeIf { it != "—" } ?: stop.telegramTime
                            val diffStr = calculateMinutesDiff(stop.afishaTime, activeComparisonTime)

                            val isTgCanceled = trainRow.hasTelegram && (stop.isSkippedByTelegram || stop.telegramTime.contains("ОТМЕН", true))
                            val isExpCanceled = stop.exportTime.contains("ОТМЕН", true)

                            Row(modifier = Modifier.fillMaxWidth().padding(horizontal = 10.dp, vertical = 4.dp), verticalAlignment = Alignment.CenterVertically) {
                                Text(stop.stationName, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.5f))
                                Text(stop.afishaTime, fontFamily = FontFamily.Monospace, fontSize = 11.sp, modifier = Modifier.weight(0.9f))
                                Text(
                                    if (trainRow.hasTelegram) stop.telegramTime else "—",
                                    fontFamily = FontFamily.Monospace,
                                    fontWeight = FontWeight.Bold,
                                    color = if (isTgCanceled) TrainRTheme.RedCancel else TrainRTheme.TextDark,
                                    fontSize = 11.sp,
                                    modifier = Modifier.weight(1.1f)
                                )
                                Text(
                                    stop.exportTime,
                                    fontFamily = FontFamily.Monospace,
                                    fontWeight = FontWeight.Bold,
                                    color = if (isExpCanceled) TrainRTheme.RedCancel else TrainRTheme.TextSub,
                                    fontSize = 11.sp,
                                    modifier = Modifier.weight(1.1f)
                                )
                                Text(diffStr, fontFamily = FontFamily.Monospace, fontSize = 11.sp, modifier = Modifier.weight(0.9f))
                                Text(
                                    when {
                                        isExpCanceled || isTgCanceled -> "БЕЗ ОСТАНОВКИ / ОТМЕНА"
                                        stop.isAddedByTelegram -> "НАЗНАЧЕНА ОСТАНОВКА"
                                        stop.isTechStop -> "ТЕХ. СТОЯНКА (без пасс.)"
                                        else -> "Посадка / высадка пасс."
                                    },
                                    fontSize = 10.sp,
                                    color = if (isExpCanceled || isTgCanceled) TrainRTheme.RedCancel else TrainRTheme.TextDark,
                                    modifier = Modifier.weight(1.5f)
                                )
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                Row(modifier = Modifier.weight(0.48f).fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(14.dp)) {
                    Column(modifier = Modifier.weight(1f).fillMaxHeight()) {
                        val tgTitle = if (trainRow.hasTelegram) {
                            "Выдержка из Телеграммы (${trainRow.telegramNumber}, от ${trainRow.telegramDate}, ${trainRow.telegramPage}):"
                        } else {
                            "Выдержка из Телеграммы:"
                        }
                        Text(tgTitle, fontWeight = FontWeight.Bold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
                        Spacer(modifier = Modifier.height(4.dp))
                        Surface(color = Color(0xFFF1F5F9), shape = RoundedCornerShape(8.dp), modifier = Modifier.fillMaxSize()) {
                            Text(
                                if (trainRow.hasTelegram && trainRow.rawTelegramSnippet.isNotBlank()) {
                                    "Номер ТГ: ${trainRow.telegramNumber}\nДата ТГ: ${trainRow.telegramDate}\nСтраница: ${trainRow.telegramPage.ifBlank { "—" }}\n\n${trainRow.rawTelegramSnippet}"
                                } else {
                                    "Телеграмма для данного поезда не загружена."
                                },
                                fontSize = 11.sp,
                                fontFamily = FontFamily.Monospace,
                                modifier = Modifier.padding(12.dp).verticalScroll(rememberScrollState())
                            )
                        }
                    }
                    Column(modifier = Modifier.weight(1f).fillMaxHeight()) {
                        Text("Выдержка из Выгрузки (${trainRow.exportBasis}):", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
                        Spacer(modifier = Modifier.height(4.dp))
                        Surface(color = Color(0xFFF1F5F9), shape = RoundedCornerShape(8.dp), modifier = Modifier.fillMaxSize()) {
                            Text(
                                trainRow.rawExportSnippet.ifBlank { "Выгрузка для данного поезда на эту дату не загружена." },
                                fontSize = 11.sp,
                                fontFamily = FontFamily.Monospace,
                                modifier = Modifier.padding(12.dp).verticalScroll(rememberScrollState())
                            )
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween) {
                    Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.BorderCol)) {
                        Text("Свернуть окно", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold)
                    }
                    Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark)) {
                        Text("Закрыть досье", color = Color.White, fontWeight = FontWeight.Bold)
                    }
                }
            }
        }
    }
}

/**
 * Календарь оперативных изменений на ВСЕ 12 МЕСЯЦЕВ ГОДА с выбором года,
 * возможностью выбирать отдельные дни, целые месяцы или все дни с изменениями.
 */
@Composable
fun CalendarGridPickerDialog(
    availableDatesWithChanges: Set<String>,
    selectedDates: Set<String>,
    onToggleDate: (String) -> Unit,
    onSelectMultipleDates: (Set<String>) -> Unit,
    onClearAll: () -> Unit,
    onDismiss: () -> Unit
) {
    var selectedYear by remember { mutableStateOf(2026) }
    var allowAnyDayClick by remember { mutableStateOf(true) }

    Dialog(onDismissRequest = onDismiss, properties = DialogProperties(usePlatformDefaultWidth = false)) {
        Surface(
            color = Color.White,
            shape = RoundedCornerShape(12.dp),
            modifier = Modifier.fillMaxWidth(0.92f).fillMaxHeight(0.88f)
        ) {
            Column(modifier = Modifier.fillMaxSize().padding(18.dp)) {
                // Верхняя панель календаря
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Column {
                        Text("КАЛЕНДАРЬ ОПЕРАТИВНЫХ ИЗМЕНЕНИЙ (ВСЕ 12 МЕСЯЦЕВ ГОДА)", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                        Text(
                            "Желтым подсвечены дни с обнаруженными изменениями. Клик по заголовку месяца выбирает дни месяца.",
                            fontSize = 11.sp,
                            color = TrainRTheme.TextSub
                        )
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp), verticalAlignment = Alignment.CenterVertically) {
                        OutlinedButton(onClick = { selectedYear-- }, contentPadding = PaddingValues(horizontal = 10.dp, vertical = 4.dp)) {
                            Text("◀ ${selectedYear - 1}", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                        }
                        Surface(color = TrainRTheme.NavyDark, shape = RoundedCornerShape(6.dp)) {
                            Text("$selectedYear год", color = Color.White, fontWeight = FontWeight.ExtraBold, fontSize = 13.sp, modifier = Modifier.padding(horizontal = 12.dp, vertical = 6.dp))
                        }
                        OutlinedButton(onClick = { selectedYear++ }, contentPadding = PaddingValues(horizontal = 10.dp, vertical = 4.dp)) {
                            Text("${selectedYear + 1} ▶", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                        }
                        Spacer(modifier = Modifier.width(8.dp))
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.BorderCol)) {
                            Text("Свернуть", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RedCancel)) {
                            Text("Закрыть", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                Spacer(modifier = Modifier.height(10.dp))

                // Сетка всех 12 месяцев (4 ряда по 3 месяца) с прокруткой
                Column(
                    modifier = Modifier.weight(1f).fillMaxWidth().verticalScroll(rememberScrollState()),
                    verticalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    val monthRows = PolygonConfig.ALL_MONTHS.chunked(3)
                    for (rowMonths in monthRows) {
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(12.dp)) {
                            for (mSpec in rowMonths) {
                                val ym = YearMonth.of(selectedYear, mSpec.monthNumber)
                                MonthGridCalendar(
                                    yearMonth = ym,
                                    monthGenitiveName = mSpec.genitiveName,
                                    monthTitle = "${mSpec.titleName} $selectedYear",
                                    availableDatesWithChanges = availableDatesWithChanges,
                                    selectedDates = selectedDates,
                                    allowAnyDayClick = allowAnyDayClick,
                                    onToggleDate = onToggleDate,
                                    onSelectMonthDates = { monthDates -> onSelectMultipleDates(monthDates) },
                                    modifier = Modifier.weight(1f)
                                )
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                Spacer(modifier = Modifier.height(10.dp))

                // Нижняя панель управления выбором
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Row(horizontalArrangement = Arrangement.spacedBy(10.dp), verticalAlignment = Alignment.CenterVertically) {
                        TextButton(onClick = onClearAll) {
                            Text("Сбросить выделение (показать все дни)", color = TrainRTheme.RedCancel, fontWeight = FontWeight.Bold)
                        }
                        if (availableDatesWithChanges.isNotEmpty()) {
                            OutlinedButton(onClick = { onSelectMultipleDates(availableDatesWithChanges) }) {
                                Text("Выбрать все дни с изменениями (${availableDatesWithChanges.size})", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }
                        }
                        Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.clickable { allowAnyDayClick = !allowAnyDayClick }) {
                            Checkbox(checked = allowAnyDayClick, onCheckedChange = { allowAnyDayClick = it })
                            Text("Разрешить выбор любых дат календаря", fontSize = 11.sp, color = TrainRTheme.TextDark)
                        }
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.BorderCol)) {
                            Text("Свернуть", color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold)
                        }
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue)) {
                            Text("Применить выбранные (${selectedDates.size})", color = Color.White, fontWeight = FontWeight.Bold)
                        }
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
    allowAnyDayClick: Boolean,
    onToggleDate: (String) -> Unit,
    onSelectMonthDates: (Set<String>) -> Unit,
    modifier: Modifier = Modifier
) {
    val daysInMonth = yearMonth.lengthOfMonth()
    val monthChangeDates = (1..daysInMonth).map { "$it $monthGenitiveName" }.filter { it in availableDatesWithChanges }.toSet()

    Surface(color = Color(0xFFF8FAFC), shape = RoundedCornerShape(8.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = modifier) {
        Column(modifier = Modifier.padding(10.dp)) {
            Row(
                modifier = Modifier.fillMaxWidth().padding(bottom = 6.dp),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Text(monthTitle, fontWeight = FontWeight.ExtraBold, fontSize = 13.sp, color = TrainRTheme.NavyDark)
                if (monthChangeDates.isNotEmpty()) {
                    Text(
                        "Выбрать (${monthChangeDates.size})",
                        fontSize = 10.sp,
                        fontWeight = FontWeight.Bold,
                        color = TrainRTheme.ActionBlue,
                        modifier = Modifier.clickable { onSelectMonthDates(monthChangeDates) }
                    )
                }
            }
            val rows = (1..daysInMonth).chunked(7)
            for (week in rows) {
                Row(modifier = Modifier.fillMaxWidth().padding(vertical = 2.dp), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                    for (day in week) {
                        val dateKey = "$day $monthGenitiveName"
                        val hasChanges = dateKey in availableDatesWithChanges
                        val isClickable = hasChanges || allowAnyDayClick
                        val isSelected = dateKey in selectedDates
                        Box(
                            contentAlignment = Alignment.Center,
                            modifier = Modifier
                                .size(32.dp)
                                .clip(RoundedCornerShape(6.dp))
                                .background(
                                    when {
                                        isSelected -> TrainRTheme.ActionBlue
                                        hasChanges -> TrainRTheme.GoldAccent.copy(alpha = 0.28f)
                                        else -> Color.White
                                    }
                                )
                                .border(
                                    1.dp,
                                    when {
                                        isSelected -> TrainRTheme.ActionBlue
                                        hasChanges -> TrainRTheme.GoldAccent
                                        else -> TrainRTheme.BorderCol
                                    },
                                    RoundedCornerShape(6.dp)
                                )
                                .let { mod -> if (isClickable) mod.clickable { onToggleDate(dateKey) } else mod }
                        ) {
                            Text(
                                day.toString(),
                                fontSize = 11.sp,
                                fontWeight = if (hasChanges || isSelected) FontWeight.ExtraBold else FontWeight.Normal,
                                color = when {
                                    isSelected -> Color.White
                                    hasChanges -> TrainRTheme.NavyDark
                                    isClickable -> TrainRTheme.TextDark
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

fun pickFilesDialog(title: String, extensions: Array<String>, multiSelect: Boolean): List<File> {
    val chooser = JFileChooser().apply {
        dialogTitle = title
        isMultiSelectionEnabled = multiSelect
        fileFilter = FileNameExtensionFilter("Поддерживаемые файлы (${extensions.joinToString(", ") { "*.$it" }})", *extensions)
    }
    val res = chooser.showOpenDialog(null)
    if (res == JFileChooser.APPROVE_OPTION) {
        return if (multiSelect) chooser.selectedFiles?.toList() ?: emptyList() else listOfNotNull(chooser.selectedFile)
    }
    return emptyList()
}
