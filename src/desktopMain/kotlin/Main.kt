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
import androidx.compose.foundation.text.BasicTextField
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
import androidx.compose.ui.graphics.SolidColor
import androidx.compose.ui.text.TextStyle
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
import java.awt.Toolkit
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.math.BigInteger
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

enum class YandexViewMode(val label: String) {
    PDF_DATA("ДАННЫЕ ИЗ PDF ЯНДЕКСА"),
    COMPARE_CHANGES("СВЕРКА С ВЫГРУЗКОЙ / ТГ"),
    RAW_TEXT("ТЕКСТ PDF ДОКУМЕНТА")
}

enum class AuditStatus(val label: String, val color: Color) {
    OK("СОВПАДАЕТ", Color(0xFF16A34A)),
    ONLY_EXPORT("ВЫГРУЗКА (БЕЗ ТГ)", Color(0xFF2563EB)),
    ONLY_TELEGRAM("ТОЛЬКО В ТГ", Color(0xFF0284C7)),
    CRITICAL("КРИТИЧЕСКАЯ ОШИБКА", Color(0xFFDC2626)),
    TIME_MISMATCH("РАСХОЖДЕНИЕ ВРЕМЕНИ", Color(0xFFD97706)),
    STOP_ERROR("ОШИБКА ОСТАНОВОК", Color(0xFF7C3AED)),
    GLOBAL_INFRA_ERROR("ЗАКРЫТАЯ СТАНЦИЯ", Color(0xFF991B1B))
}

data class TelegramExcerptItem(
    val telegramNumber: String,
    val telegramDate: String,
    val telegramPage: String,
    val fileName: String,
    val datesText: String = "",
    val snippet: String
)

data class StationStopEntry(
    val stationName: String,
    val afishaTime: String = "—",
    val telegramTime: String = "—",
    val exportTime: String = "—",
    val yandexTime: String = "—",
    val isInitialStation: Boolean = false,
    val isFinalStation: Boolean = false,
    val isFullTrainCanceled: Boolean = false,
    val isSkippedByTelegram: Boolean = false,
    val isSkippedByExport: Boolean = false,
    val isAddedByTelegram: Boolean = false,
    val isAddedByExport: Boolean = false,
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
    val hasExport: Boolean = false,
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
    val telegramExcerpts: List<TelegramExcerptItem> = emptyList(),
    val rawExportSnippet: String = "",
    val fullThread: List<StationStopEntry> = emptyList(),
    val telegramThread: List<StationStopEntry> = emptyList(),
    val auditStatus: AuditStatus = AuditStatus.OK
)

data class AfishaBaseTrain(
    val trainNumber: String,
    val route: String,
    val isToMoscow: Boolean,
    val stationTimes: Map<String, String>
)

data class YandexParsedTrain(
    val id: String,
    val trainNumber: String,
    val dateStr: String,
    val station: String,
    val departureTime: String,
    val arrivalTime: String = "—",
    val routeAndDetails: String = "",
    val statusText: String,
    val sourceFileName: String = "",
    val pageNumber: Int = 1
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

data class TelegramParseResult(
    val stationMap: StationDateTrainMap,
    val excerptsByTrain: Map<String, List<TelegramExcerptItem>>,
    val threadsByTrain: Map<String, List<StationStopEntry>>
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
            .replace(Regex("(?iu)^(?:ст\\.|станция|оп|о\\.п\\.|пл\\.|платформа)\\s*"), "")
            .replace(Regex("(?iu)Москва[\\s\\-]*Пасс\\.?[\\s\\-]*(?:Окт(?:ябрьская)?\\.?)?"), "Москва")
            .replace(Regex("(?iu)Москва[\\s\\-]*Пассажирская(?:[\\s\\-]*Октябрьская)?"), "Москва")
            .replace(Regex("(?iu)Москва[\\s\\-]*Окт(?:ябрьская)?\\.?"), "Москва")
            .replace(Regex("(?iu)Москва\\s*\\(Ленинградский\\s*вокзал\\)"), "Москва")
            .replace(Regex("(?iu)Москва[\\s\\-]*Тов\\.?[\\s\\-]*(?:Окт(?:ябрьская)?\\.?)?"), "Москва-Товарная")
            .replace(Regex("(?iu)Крюково\\s*\\(Зеленоград\\)"), "Крюково")
            .replace("Зеленоград-Крюково", "Крюково", ignoreCase = true)
            .replace("Грачевская", "Грачёвская", ignoreCase = true)
            .replace("Березки-Дачные", "Берёзки-Дачные", ignoreCase = true)
            .replace("Березки Дачные", "Берёзки-Дачные", ignoreCase = true)
            .replace("Поварово 1", "Поварово-1", ignoreCase = true)
            .replace("Поварово-I", "Поварово-1", ignoreCase = true)
            .replace("Ленинградский вокзал", "Москва", ignoreCase = true)
            .replace("Оп Подрезково", "Подрезково", ignoreCase = true)
            .replace("Бологое-Моск.", "Бологое", ignoreCase = true)
            .replace("Конаково-ГРЭС", "Конаково ГРЭС", ignoreCase = true)
            .trim()
        return ALL_STATIONS.firstOrNull { it.equals(clean, ignoreCase = true) } ?: clean
    }

    fun isTrainToMoscow(trainNumber: String, route: String = ""): Boolean {
        val cleanRoute = route.replace(Regex("\\([^)]*\\)"), "")
        val r = cleanRoute.lowercase(Locale.getDefault())
        if (r.contains("на москву") || r.contains("до москвы") || r.contains("до ипподрома") ||
            r.contains("на ипподром") || r.contains("до 47 км") || r.contains("-москва") || r.contains("— москва")
        ) {
            return true
        }
        if (r.contains("из москвы") || r.contains("от москвы") || r.contains("из ипподрома") ||
            r.contains("до твери") || r.contains("до клина") || r.contains("до крюково") ||
            r.contains("до решетниково") || r.contains("до конаково") || r.contains("до подсолнечной") ||
            r.contains("до сходни") || r.contains("до завидово") || r.contains("до бологое") ||
            r.startsWith("москва-") || r.startsWith("москва —")
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
// 3. ТОЧКА ВХОДА ПРИЛОЖЕНИЯ
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
            TrainRApp()
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
        else -> "ноября"
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

fun extractSmartDates(rawText: String): List<String> {
    val text = rawText.replace(Regex("(?iu)\\bот\\s+\\d{1,2}[./]\\d{2}(?:[./]\\d{2,4})?"), " ")
    val result = mutableListOf<String>()
    val monthNamesPattern = "(январ[яе]|феврал[яе]|март[ае]?|апрел[яе]|ма[яе]|июн[яе]|июл[яе]|август[ае]?|сентябр[яе]|октябр[яе]|ноябр[яе]|декабр[яе])"

    val slashRangeRegex = Regex("(?iu)(?:с\\s+)?(\\d{1,2})[./](\\d{2})\\s*(?:по|[-–—])\\s*(\\d{1,2})[./](\\d{2})")
    for (m in slashRangeRegex.findAll(text)) {
        val d1 = m.groupValues[1].toIntOrNull() ?: continue
        val m1 = m.groupValues[2].toIntOrNull() ?: continue
        val d2 = m.groupValues[3].toIntOrNull() ?: continue
        val m2 = m.groupValues[4].toIntOrNull() ?: continue
        if (m1 == m2 && m1 in 1..12 && d1 in 1..31 && d2 in d1..31) {
            val monthGen = getMonthNameGenitive(m.groupValues[2])
            for (d in d1..d2) result.add("$d $monthGen")
        }
    }

    val slashGroupRegex = Regex("(?iu)(?<!\\d)(\\d{1,2}(?:\\s*(?:,|[-–—]|и)\\s*\\d{1,2})*)[./](\\d{2})(?:[./]\\d{2,4})?(?![\\d/])")
    for (m in slashGroupRegex.findAll(text)) {
        val monthNum = m.groupValues[2].toIntOrNull() ?: continue
        if (monthNum !in 1..12) continue
        val monthGen = getMonthNameGenitive(m.groupValues[2])
        val daysExpr = m.groupValues[1]
        for (part in daysExpr.split(Regex("(?iu),|\\s+и\\s+"))) {
            val trimmed = part.trim()
            val rangeMatch = Regex("^(\\d{1,2})\\s*[-–—]\\s*(\\d{1,2})$").find(trimmed)
            if (rangeMatch != null) {
                val d1 = rangeMatch.groupValues[1].toIntOrNull() ?: continue
                val d2 = rangeMatch.groupValues[2].toIntOrNull() ?: continue
                if (d1 in 1..31 && d2 in d1..31) {
                    for (d in d1..d2) result.add("$d $monthGen")
                }
            } else {
                val d = trimmed.toIntOrNull()
                if (d != null && d in 1..31) result.add("$d $monthGen")
            }
        }
    }

    val wordGroupRegex = Regex("(?iu)((?:(?:с\\s+)?\\d{1,2}(?:\\s*(?:по|[-–—]|,|и)\\s*\\d{1,2})*)+)\\s+$monthNamesPattern")
    for (m in wordGroupRegex.findAll(text)) {
        val daysExpr = m.groupValues[1].replace(Regex("(?iu)\\bс\\b"), " ").replace(Regex("(?iu)\\bпо\\b"), "-")
        val monthGen = normalizeMonthGenitive(m.groupValues[2])
        for (part in daysExpr.split(Regex("(?iu),|\\s+и\\s+"))) {
            val trimmed = part.trim()
            val rangeMatch = Regex("^(\\d{1,2})\\s*[-–—]\\s*(\\d{1,2})$").find(trimmed)
            if (rangeMatch != null) {
                val d1 = rangeMatch.groupValues[1].toIntOrNull() ?: continue
                val d2 = rangeMatch.groupValues[2].toIntOrNull() ?: continue
                if (d1 in 1..31 && d2 in d1..31) {
                    for (d in d1..d2) result.add("$d $monthGen")
                }
            } else {
                val d = trimmed.toIntOrNull()
                if (d != null && d in 1..31) result.add("$d $monthGen")
            }
        }
    }

    return result.distinct()
}

fun extractSmartRoute(text: String, trainNum: String, isToMoscow: Boolean, afishaTrain: AfishaBaseTrain?): String {
    val textWithoutParens = text.replace(Regex("\\([^)]*\\)"), " ")

    val tgHeaderRouteRegex = Regex("(?iu)(?:нр\\.|№)\\s*[1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?\\s+([А-Яа-яЁё\\-]+)\\s*[-–—]\\s*([А-Яа-яЁё\\-]+(?:\\s+[А-Яа-яЁё.]+)?)")
    val tgHeaderMatch = tgHeaderRouteRegex.find(textWithoutParens)
    if (tgHeaderMatch != null) {
        val fromSt = PolygonConfig.normalizeStationName(tgHeaderMatch.groupValues[1])
        val toSt = PolygonConfig.normalizeStationName(tgHeaderMatch.groupValues[2])
        if (toSt.equals("Москва", ignoreCase = true)) return "до Москвы"
        if (toSt.contains("Ипподром", ignoreCase = true)) return "до Ипподрома (47 км)"
        return "$fromSt — $toSt"
    }

    val exportRouteRegex = Regex("(?iu)№\\s*[1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?\\s+((?:из\\s+[А-Яа-яЁё\\-]+\\s+)?до\\s+[А-Яа-яЁё\\-]+(?:\\s*\\(\\d+\\s*км\\))?)")
    val expMatch = exportRouteRegex.find(textWithoutParens)
    if (expMatch != null) {
        return expMatch.groupValues[1].trim().replace(Regex("\\s{2,}"), " ")
    }

    if (textWithoutParens.contains("Ипподром", ignoreCase = true) || textWithoutParens.contains("47 км", ignoreCase = true)) {
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
    val matches = Regex("(?iu)$pattern").findAll(text)
    val result = mutableListOf<String>()
    for (match in matches) {
        val listStr = match.groupValues[1]
        val parts = listStr.split(Regex("(?iu),|;|\\s+и\\s+"))
        for (p in parts) {
            val cleaned = p.replace(".", "").trim()
            val norm = PolygonConfig.normalizeStationName(cleaned)
            if (norm in PolygonConfig.ALL_STATIONS) {
                result.add(norm)
            } else {
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

    val cleanPath = pathChanges.replace(Regex("(?iu)(?:^|[^а-яА-ЯёЁ])Изменений\\s*нет\\.?"), "").trim()
    if (cleanPath.isNotEmpty()) {
        parts.add(cleanPath)
    }

    if (parts.isEmpty()) {
        return if (forWord) "" else "по графику со всеми остановками"
    }
    return parts.joinToString(". ").replace(Regex("\\.{2,}"), ".")
}

fun calculateMinutesDiff(timeA: String, timeB: String): String {
    if (timeB.contains("ОТМЕН", true) || timeB.contains("БЕЗ ОСТ", true) || timeB.contains("ТЕХ", true)) return "ОТМЕНА"
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

    val adjustedHours = if (hours in 0..3) hours + 24 else hours
    return adjustedHours * 60 + minutes
}

fun getSortingTime(row: TrainChangeRow): Int {
    val baseRegex = Regex("(?iu)(?:отправлением\\s+в|назначается[^.\\n]*?\\s+в)\\s+(\\d{1,2}[:.]\\d{2})")
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

fun mergeSameTrainRowsForDay(
    rows: List<TrainChangeRow>,
    globalExcerptsByTrain: Map<String, List<TelegramExcerptItem>> = emptyMap(),
    globalThreadsByTrain: Map<String, List<StationStopEntry>> = emptyMap()
): List<TrainChangeRow> {
    if (rows.isEmpty()) return emptyList()
    val grouped = rows.groupBy { "${it.dateStr}|${it.station}|${it.trainNumber}" }
    val mergedList = mutableListOf<TrainChangeRow>()

    for ((_, sameTrains) in grouped) {
        val tNum = sameTrains.first().trainNumber
        val tgRows = sameTrains.filter { it.hasTelegram }
        val expRows = sameTrains.filter { it.hasExport }
        val base = expRows.firstOrNull() ?: sameTrains.first()

        val globalExcerpts = globalExcerptsByTrain[tNum] ?: emptyList()
        val rowExcerpts = sameTrains.flatMap { it.telegramExcerpts }
        val allDistinctExcerpts = (rowExcerpts + globalExcerpts)
            .distinctBy { "${it.fileName}|${it.telegramPage}|${it.snippet.take(120)}" }

        val hasTg = tgRows.isNotEmpty() || allDistinctExcerpts.isNotEmpty()
        val hasExp = expRows.isNotEmpty()

        val combinedSkipped = sameTrains.flatMap { it.skippedStops }.distinct()
        val combinedAdded = sameTrains.flatMap { it.addedStops }.distinct()

        val combinedFullCancel = if (hasExp) expRows.any { it.isFullCancel } else sameTrains.any { it.isFullCancel }
        val combinedPartial = sameTrains.any { it.isPartialCancel }
        val combinedSkippedAtStation = if (hasExp) expRows.any { it.isSkippedAtStation } else sameTrains.any { it.isSkippedAtStation }
        val combinedCanceled = combinedFullCancel || combinedSkippedAtStation
        val combinedAppointed = sameTrains.any { it.isAppointed }

        val finalSkipped = if (combinedFullCancel) emptyList() else combinedSkipped

        val latestTelegramTime = if (tgRows.isNotEmpty()) {
            tgRows.map { it.telegramTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: "—"
        } else {
            globalThreadsByTrain[tNum]?.firstOrNull { it.stationName.equals(base.station, true) }?.telegramTime ?: "—"
        }

        val latestExportTime = if (hasExp) {
            expRows.map { it.exportTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: "—"
        } else {
            "—"
        }

        val latestYandexTime = sameTrains.map { it.yandexTime }.lastOrNull { it != "—" && it.isNotBlank() } ?: "—"

        val bestAfishaTime = expRows.map { it.afishaTime }.firstOrNull { it != "—" && it.isNotBlank() }
            ?: sameTrains.map { it.afishaTime }.firstOrNull { it != "—" && it.isNotBlank() }
            ?: "—"

        val bestRoute = expRows.map { it.route }.firstOrNull { it.isNotBlank() }
            ?: tgRows.map { it.route }.firstOrNull { it.isNotBlank() }
            ?: base.route

        val mergedActualText = expRows.map { it.actualChangesText.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: sameTrains.map { it.actualChangesText.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: ""

        val mergedRawTg = allDistinctExcerpts.mapIndexed { idx, item ->
            "=== Телеграмма #${idx + 1}: ${item.telegramNumber} от ${item.telegramDate} (${item.telegramPage}, файл: ${item.fileName}) ===\n${item.snippet}"
        }.joinToString("\n\n")

        val mergedRawExp = expRows.map { it.rawExportSnippet.trim() }
            .filter { it.isNotBlank() }
            .distinct()
            .joinToString("\n")

        val mergedWordLine = expRows.map { it.wordExportLine.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: sameTrains.map { it.wordExportLine.trim() }.filter { it.isNotEmpty() }.maxByOrNull { it.length }
            ?: ""

        val mergedTgNumber = allDistinctExcerpts.map { it.telegramNumber.trim() }.filter { it.isNotEmpty() }.distinct().joinToString(" | ")
        val mergedTgDate = allDistinctExcerpts.map { it.telegramDate.trim() }.filter { it.isNotEmpty() }.distinct().joinToString(" | ")
        val mergedTgPage = allDistinctExcerpts.map { it.telegramPage.trim() }.filter { it.isNotEmpty() }.distinct().joinToString(", ")

        val mergedTgBasis = if (hasTg) {
            allDistinctExcerpts.map { "${it.telegramNumber} от ${it.telegramDate} (${it.telegramPage})" }.distinct().joinToString("; ")
                .ifEmpty { "Телеграмма загружена" }
        } else {
            "—"
        }

        val mergedExpBasis = expRows.map { it.exportBasis.trim() }.filter { it.isNotEmpty() && it != "—" }.distinct().joinToString(", ")
            .ifEmpty { "—" }

        val mergedError = sameTrains.mapNotNull { it.globalErrorNote }.distinct().joinToString("; ").ifEmpty { null }
        val bestTgThread = tgRows.map { it.telegramThread }.firstOrNull { it.isNotEmpty() }
            ?: globalThreadsByTrain[tNum]
            ?: emptyList()

        val evaluatedStatus = evaluateFourWayAuditStatus(
            afishaTime = bestAfishaTime,
            telegramTime = latestTelegramTime,
            exportTime = latestExportTime,
            yandexTime = latestYandexTime,
            hasTelegram = hasTg,
            hasExport = hasExp,
            isCanceled = combinedCanceled,
            skippedStops = finalSkipped,
            globalError = mergedError
        )

        mergedList.add(
            base.copy(
                route = bestRoute,
                hasTelegram = hasTg,
                hasExport = hasExp,
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
                telegramExcerpts = allDistinctExcerpts,
                rawExportSnippet = mergedRawExp,
                telegramThread = bestTgThread,
                auditStatus = evaluatedStatus
            )
        )
    }
    return mergedList.sortedBy { getSortingTime(it) }
}

fun buildFullThreads(
    map: StationDateTrainMap,
    afishaMap: Map<String, AfishaBaseTrain>,
    yandexList: List<YandexParsedTrain> = emptyList()
): StationDateTrainMap {
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
            val tgThreadTemplate = rows.firstNotNullOfOrNull { it.telegramThread.takeIf { th -> th.isNotEmpty() } } ?: emptyList()
            val tgByStation = tgThreadTemplate.associateBy { it.stationName }

            val isTrainFullCanceledInExport = rows.any { it.hasExport && it.isFullCancel }
            val isTrainFullCanceledInTg = rows.any { it.hasTelegram && it.isFullCancel }
            val anyFullCanceled = isTrainFullCanceledInExport || isTrainFullCanceledInTg

            val allExportSkippedStops = rows.filter { it.hasExport }.flatMap { it.skippedStops }.toSet()
            val allExportAddedStops = rows.filter { it.hasExport }.flatMap { it.addedStops }.toSet()

            val parsedStations = rows.associateBy { it.station }
            val afishaStations = baseTrain?.stationTimes?.keys ?: emptySet()
            val tgStations = tgByStation.keys

            val allStations = (parsedStations.keys + afishaStations + tgStations).distinct()

            val sortedStations = allStations.sortedBy { st ->
                val idx = PolygonConfig.ALL_STATIONS.indexOfFirst { it.equals(st, ignoreCase = true) }
                if (idx == -1) 999 else if (isToMsk) -idx else idx
            }

            val threadEntries = sortedStations.map { st ->
                val parsedRow = parsedStations[st]
                val tgStop = tgByStation[st]
                val afishaT = baseTrain?.stationTimes?.get(st)
                    ?: parsedRow?.afishaTime?.takeIf { it != "—" && parsedRow.hasExport }
                    ?: "—"

                val isSkippedByExpAtSt = st in allExportSkippedStops ||
                    (parsedRow != null && parsedRow.hasExport && (parsedRow.isCanceled || parsedRow.exportTime.contains("ОТМЕН", true)))

                val expT = when {
                    parsedRow != null && parsedRow.hasExport -> parsedRow.exportTime
                    isTrainFullCanceledInExport -> if (afishaT != "—") "$afishaT ОТМЕНА" else "ОТМЕНА"
                    st in allExportSkippedStops -> if (afishaT != "—") "$afishaT ОТМЕНА" else "БЕЗ ОСТ."
                    else -> "—"
                }

                val yandexMatch = yandexList.firstOrNull {
                    (it.trainNumber == tNum || (it.trainNumber == "—" && (it.departureTime == afishaT || it.departureTime == expT.take(5)))) &&
                        (it.station.equals(st, true) || it.station.isBlank()) &&
                        (it.dateStr == dt || it.dateStr == "Все даты")
                }
                val yndT = yandexMatch?.let {
                    if (it.statusText.contains("Отмен", true)) "${it.departureTime} ОТМЕНА" else it.departureTime
                } ?: parsedRow?.yandexTime ?: "—"

                if (tgStop != null) {
                    tgStop.copy(
                        afishaTime = if (afishaT != "—") afishaT else tgStop.afishaTime,
                        exportTime = expT,
                        yandexTime = yndT,
                        isFullTrainCanceled = anyFullCanceled,
                        isSkippedByExport = isSkippedByExpAtSt || isTrainFullCanceledInExport,
                        isAddedByExport = st in allExportAddedStops
                    )
                } else if (parsedRow != null) {
                    val isSkippedTg = parsedRow.hasTelegram && (
                        parsedRow.isCanceled ||
                            parsedRow.telegramTime.contains("ОТМЕН", true) ||
                            parsedRow.telegramTime.contains("БЕЗ ОСТ", true)
                        )

                    StationStopEntry(
                        stationName = st,
                        afishaTime = afishaT,
                        telegramTime = if (parsedRow.hasTelegram) parsedRow.telegramTime else "—",
                        exportTime = expT,
                        yandexTime = yndT,
                        isFullTrainCanceled = anyFullCanceled || parsedRow.isFullCancel,
                        isSkippedByTelegram = isSkippedTg,
                        isSkippedByExport = isSkippedByExpAtSt || isTrainFullCanceledInExport,
                        isAddedByTelegram = parsedRow.hasTelegram && parsedRow.isAppointed,
                        isAddedByExport = st in allExportAddedStops || (parsedRow.hasExport && parsedRow.isAppointed),
                        isTechStop = false
                    )
                } else {
                    StationStopEntry(
                        stationName = st,
                        afishaTime = afishaT,
                        telegramTime = if (isTrainFullCanceledInTg) "ОТМЕНА" else "—",
                        exportTime = expT,
                        yandexTime = yndT,
                        isFullTrainCanceled = anyFullCanceled,
                        isSkippedByTelegram = isTrainFullCanceledInTg,
                        isSkippedByExport = isSkippedByExpAtSt || isTrainFullCanceledInExport,
                        isAddedByExport = st in allExportAddedStops
                    )
                }
            }

            for (r in rows) {
                val cleanExpDep = Regex("(\\d{1,2}:\\d{2})").find(r.exportTime)?.groupValues?.get(1) ?: ""
                val matchedYandex = yandexList.firstOrNull {
                    (it.trainNumber == tNum || (it.trainNumber == "—" && (it.departureTime == r.afishaTime || it.departureTime == cleanExpDep))) &&
                        (it.station.equals(r.station, true) || it.station.isBlank()) &&
                        (it.dateStr == dt || it.dateStr == "Все даты")
                }
                val updatedYandexTime = matchedYandex?.let {
                    if (it.statusText.contains("Отмен", true)) "${it.departureTime} ОТМЕНА" else it.departureTime
                } ?: r.yandexTime

                val updatedStatus = evaluateFourWayAuditStatus(
                    afishaTime = r.afishaTime,
                    telegramTime = r.telegramTime,
                    exportTime = r.exportTime,
                    yandexTime = updatedYandexTime,
                    hasTelegram = r.hasTelegram,
                    hasExport = r.hasExport,
                    isCanceled = r.isCanceled,
                    skippedStops = r.skippedStops,
                    globalError = r.globalErrorNote
                )

                resultMap.getOrPut(r.station) { mutableMapOf() }
                    .getOrPut(dt) { mutableMapOf() }[tNum] = r.copy(
                    yandexTime = updatedYandexTime,
                    auditStatus = updatedStatus,
                    fullThread = threadEntries
                )
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
        val tgCanceled = telegramTime.contains("ОТМЕН", true) || telegramTime.contains("БЕЗ ОСТ", true) || telegramTime.contains("ТЕХ.", true)
        val expCanceled = exportTime.contains("ОТМЕН", true)
        if (tgCanceled != expCanceled) return AuditStatus.CRITICAL
        val cleanTgTime = Regex("(\\d{1,2}:\\d{2})").find(telegramTime)?.groupValues?.get(1) ?: telegramTime
        val cleanExpTime = Regex("(\\d{1,2}:\\d{2})").find(exportTime)?.groupValues?.get(1) ?: exportTime
        if (!tgCanceled && cleanTgTime != cleanExpTime) return AuditStatus.TIME_MISMATCH
    }
    if (yandexTime != "—") {
        val refTime = if (hasExport && exportTime != "—") exportTime else telegramTime
        val cleanRef = Regex("(\\d{1,2}:\\d{2})").find(refTime)?.groupValues?.get(1) ?: refTime
        val cleanYnd = Regex("(\\d{1,2}:\\d{2})").find(yandexTime)?.groupValues?.get(1) ?: yandexTime
        if (refTime != "—" && !isCanceled && cleanRef != cleanYnd) {
            return AuditStatus.TIME_MISMATCH
        }
    }
    if (skippedStops.isNotEmpty() && yandexTime != "—" && yandexTime == afishaTime) {
        return AuditStatus.STOP_ERROR
    }
    if (!hasTelegram && hasExport) return AuditStatus.ONLY_EXPORT
    if (hasTelegram && !hasExport) return AuditStatus.ONLY_TELEGRAM
    return AuditStatus.OK
}

// ============================================================================
// 5. УМНЫЙ ПАРСИНГ ФАЙЛОВ
// ============================================================================

object FileParsers {
    private val cacheDir = File(System.getProperty("user.home"), ".train_r_system")
    val cachedAfishaFile = File(cacheDir, "cached_afisha.xlsx")

    fun ensureCacheAfisha(sourceFile: File): File {
        if (!cacheDir.exists()) cacheDir.mkdirs()
        sourceFile.copyTo(cachedAfishaFile, overwrite = true)
        return cachedAfishaFile
    }

    fun readSmartTextFile(file: File, sortPdfByPosition: Boolean = false): String {
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
                val stripper = PDFTextStripper().apply {
                    sortByPosition = sortPdfByPosition
                }
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

    fun readPagedTextChunks(file: File, sortPdfByPosition: Boolean = false): List<PageTextChunk> {
        val ext = file.extension.lowercase(Locale.getDefault())
        if (ext == "pdf") {
            val pages = mutableListOf<PageTextChunk>()
            Loader.loadPDF(file).use { pdDoc ->
                val total = pdDoc.numberOfPages
                val stripper = PDFTextStripper().apply {
                    sortByPosition = sortPdfByPosition
                }
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

        val fullText = readSmartTextFile(file, sortPdfByPosition)
        if (fullText.contains("\u000C")) {
            return fullText.split("\u000C").mapIndexedNotNull { idx, part ->
                if (part.isNotBlank()) PageTextChunk(idx + 1, part) else null
            }
        }

        val chunks = mutableListOf<PageTextChunk>()
        var currentPage = 1
        val sb = StringBuilder()
        val pageMarkerRegex = Regex("(?iu)^\\s*(?:стр\\.?|страница|лист)\\s*(\\d{1,3})\\b")
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

        val numRegexes = listOf(
            Regex("(?iu)(?:телеграмма|тг)\\s*(?:ЦППК|МТППК|ОКТ|РЖД)?\\s*(?:от\\s*[^№\\n]+)?\\s*(?:№|НР|N)\\s*([\\d/\\-А-Яа-яA-Za-z]+)"),
            Regex("(?iu)(?:НР|№)\\s*(\\d{1,5}[/\\-][\\d/\\-А-Яа-яA-Za-z]+)"),
            Regex("(?iu)телеграмма\\s*(?:№|НР)?\\s*([\\d/\\-А-Яа-яA-Za-z]{2,})")
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

        val dateRegexes = listOf(
            Regex("(?iu)от\\s+(\\d{1,2}\\.\\d{2}\\.\\d{2,4}(?:\\s*г\\.?)?)"),
            Regex("(?iu)от\\s+(\\d{1,2}\\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*\\d{0,4}\\s*(?:г\\.?)?)"),
            Regex("(?iu)\\b(\\d{2}\\.\\d{2}\\.\\d{2,4})\\b")
        )
        var tgDate = ""
        for (r in dateRegexes) {
            val m = r.find(headerArea) ?: r.find(file.name)
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

    fun parseTelegramFiles(files: List<File>, afishaMap: Map<String, AfishaBaseTrain>): TelegramParseResult {
        if (files.isEmpty()) return TelegramParseResult(emptyMap(), emptyMap(), emptyMap())

        val rawRowsByStationDate = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        val globalExcerptsByTrain = mutableMapOf<String, MutableList<TelegramExcerptItem>>()
        val globalThreadsByTrain = mutableMapOf<String, List<StationStopEntry>>()

        val trainRegex = Regex("(?iu)(?:нр\\.|№|поезд\\s*№?)\\s*([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)")

        val stationScheduleItemRegex = Regex(
            "(?iu)([А-Яа-яЁё][А-Яа-яЁё\\-\\s.0-9]{1,28}?)\\s+(\\d{1,2}[.:]\\d{2}(?:[.:]\\d{2})?)" +
                "(?:\\s*[-–—]\\s*(\\d{1,2}[.:]\\d{2}(?:[.:]\\d{2})?))?" +
                "(\\s*,?\\s*(?:далее\\s+)?(?:на\\s+тех\\.?\\s*стоянку|тех\\.?\\s*стоянка|стоянка\\s+техническая|без\\s+посадки[^,;)]*|без\\s+пассажиров)[^,;)]*(?:,\\s*отпр\\.?\\s*в?\\s*(\\d{1,2}[.:]\\d{2}(?:[.:]\\d{2})?))?)?"
        )

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

                    val isNewTrainLine = Regex("(?iu)^(?:\\d+[.)]\\s*)?(?:нр\\.|№|поезд)\\s*[1-7]\\d{2,3}").containsMatchIn(trimmed)
                    val isMajorHeader = Regex("(?iu)^(?:\\d+\\.\\d*\\.?\\s*)?(?:Отменить|Отменяются|Изменить|Назначить|Назначаются|Изменяется)").containsMatchIn(trimmed) &&
                        !Regex("(?iu)(?:нр\\.|№)\\s*[1-7]\\d{2,3}").containsMatchIn(trimmed)

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
            var sectionDates = extractSmartDates(file.name).toMutableList()
            var lastSectionHeader = ""

            for (pagedPara in paragraphs) {
                val rawParaText = pagedPara.text.trim()
                val cleanPara = rawParaText.replace("\n", " ").replace("\r", "")
                    .replace(Regex("(?iu)\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)"), "")
                    .replace(Regex("\\s{2,}"), " ")
                    .trim()
                if (cleanPara.isEmpty()) continue

                val trainMatches = trainRegex.findAll(cleanPara).map { it.groupValues[1] }.distinct().toList()
                val paraDates = extractSmartDates(cleanPara)

                if (trainMatches.isEmpty()) {
                    if (paraDates.isNotEmpty()) {
                        sectionDates = paraDates.toMutableList()
                    }
                    if (Regex("(?iu)(?:^|[^а-яА-ЯёЁ])(отменить|отменяются)(?:[^а-яА-ЯёЁ]|$)").containsMatchIn(cleanPara) &&
                        !Regex("(?iu)на\\s+участке|остановк").containsMatchIn(cleanPara)
                    ) {
                        isGlobalCancelMode = true
                        lastSectionHeader = cleanPara
                    } else if (Regex("(?iu)(?:^|[^а-яА-ЯёЁ])(изменить|назначить|назначаются|изменяется)(?:[^а-яА-ЯёЁ]|$)").containsMatchIn(cleanPara)) {
                        isGlobalCancelMode = false
                        lastSectionHeader = cleanPara
                    }
                    continue
                }

                val effectiveDates = when {
                    paraDates.isNotEmpty() -> paraDates
                    sectionDates.isNotEmpty() -> sectionDates
                    else -> listOf("Без даты")
                }

                val pageLabel = "Стр. ${pagedPara.pageNumber}"
                val fullExcerpt = if (lastSectionHeader.isNotBlank() && !rawParaText.startsWith(lastSectionHeader.take(15))) {
                    "[$lastSectionHeader]\n$rawParaText"
                } else {
                    rawParaText
                }

                val excerptItem = TelegramExcerptItem(
                    telegramNumber = tgNumber,
                    telegramDate = tgDate,
                    telegramPage = pageLabel,
                    fileName = file.name,
                    datesText = effectiveDates.joinToString(", "),
                    snippet = fullExcerpt
                )

                for (tNum in trainMatches) {
                    globalExcerptsByTrain.getOrPut(tNum) { mutableListOf() }.add(excerptItem)
                }

                val hasPartialRouteCancel = Regex("(?iu)(отмен[а-я]*\\s+на\\s+участке|на\\s+участке\\s+[^.;]+\\s+отмен|следует\\s+до\\s+[А-Яа-яЁё\\-]+\\s+вместо|укорочен)").containsMatchIn(cleanPara)
                val explicitFullCancelWord = Regex("(?iu)(?:^|[^а-яА-ЯёЁ])(отменяется|отмен[её]н|отменить|полная\\s+отмена)(?:[^а-яА-ЯёЁ]|$)").containsMatchIn(cleanPara) &&
                    !Regex("(?iu)отменен[а-я]*\\s*остановк").containsMatchIn(cleanPara) &&
                    !Regex("(?iu)без\\s+остановк").containsMatchIn(cleanPara) &&
                    !hasPartialRouteCancel

                val isFullCancel = (isGlobalCancelMode && !hasPartialRouteCancel && !Regex("(?iu)измененным\\s+расписанием|отправится\\s+в|назначен[ыа-я]*\\s+остановк|\\d{1,2}:\\d{2}\\s*-\\s*\\d{1,2}:\\d{2}").containsMatchIn(cleanPara)) || explicitFullCancelWord
                val isAppointed = Regex("(?iu)назначается|назначить|назначен\\s+поезд").containsMatchIn(cleanPara)

                val closedStationError = PolygonConfig.HARD_CLOSED_STATIONS.firstOrNull {
                    cleanPara.contains(it, true) && !cleanPara.contains("без остановки", true)
                }?.let { "ОШИБКА ТГ: Остановка на закрытой $it!" }

                val pathRegex = Regex("(?iu)(\\([^)]*пут[иь][^)]*\\)|по\\s*\\d+\\s*пут[иь]|проследует\\s*по\\s*\\d+\\s*пут[иь][^.,;()]*)")
                val pathStr = pathRegex.findAll(cleanPara).map { it.value }.distinct().joinToString(" ")

                val skippedStops = mutableListOf<String>()
                val addedStops = mutableListOf<String>()

                skippedStops.addAll(extractStationsByRegex(cleanPara, "отменен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)"))
                skippedStops.addAll(extractStationsByRegex(cleanPara, "без\\s*останов[ок]+\\s*:?\\s*([^.;(]+)"))
                addedStops.addAll(extractStationsByRegex(cleanPara, "назначен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)"))

                val sampleTrainForDir = trainMatches.firstOrNull() ?: ""
                val isToMskSample = PolygonConfig.isTrainToMoscow(sampleTrainForDir, cleanPara)

                data class ParsedTgStopRaw(
                    val stationName: String,
                    val t1: String,
                    val t2: String,
                    val hasTechOrNoPassNote: Boolean,
                    val techDepTime: String
                )

                val rawStopItems = mutableListOf<ParsedTgStopRaw>()
                for (sm in stationScheduleItemRegex.findAll(cleanPara)) {
                    val rawSt = sm.groupValues[1]
                        .replace(Regex("(?iu)^(?:ст\\.|оп|пл\\.|от|до|из)\\s+"), "")
                        .trim()
                    val normSt = PolygonConfig.normalizeStationName(rawSt)
                    if (normSt !in PolygonConfig.ALL_STATIONS) continue

                    val t1 = sm.groupValues[2].replace('.', ':').take(5).padStart(5, '0')
                    val t2 = sm.groupValues[3].replace('.', ':').take(5).let { if (it.isNotEmpty()) it.padStart(5, '0') else "" }
                    val techNoteText = sm.groupValues[4]
                    val techDep = sm.groupValues[5].replace('.', ':').take(5).let { if (it.isNotEmpty()) it.padStart(5, '0') else "" }

                    val hasTechOrNoPass = techNoteText.isNotBlank() &&
                        Regex("(?iu)(тех\\.?\\s*стоянк|стоянка\\s+техническая|без\\s+посадки|без\\s+высадки|без\\s+пассажиров)").containsMatchIn(techNoteText)

                    rawStopItems.add(
                        ParsedTgStopRaw(
                            stationName = normSt,
                            t1 = t1,
                            t2 = t2,
                            hasTechOrNoPassNote = hasTechOrNoPass,
                            techDepTime = techDep
                        )
                    )
                }

                val stationDepTimes = mutableMapOf<String, String>()
                val techStopStations = mutableSetOf<String>()
                val builtTgThread = mutableListOf<StationStopEntry>()

                for ((idx, item) in rawStopItems.withIndex()) {
                    val stName = item.stationName
                    val isInitial = idx == 0 && rawStopItems.size > 1
                    val isFinal = idx == rawStopItems.lastIndex && rawStopItems.size > 1
                    val isSingleStationMention = rawStopItems.size == 1

                    val isClosedSt = PolygonConfig.HARD_CLOSED_STATIONS.contains(stName) ||
                        (stName == "Левобережная" && !isToMskSample)

                    if (isClarification) {
                        if (cleanPara.contains("посадка", true)) {
                            addedStops.add(stName)
                            stationDepTimes[stName] = item.t2.ifEmpty { item.t1 }
                        } else if (cleanPara.contains("техническая", true) || cleanPara.contains("без посадки", true)) {
                            if (!isClosedSt) skippedStops.add(stName)
                            techStopStations.add(stName)
                        }
                    } else if (isInitial || isFinal) {
                        val effectiveTime = item.t2.ifEmpty { item.t1 }
                        stationDepTimes[stName] = effectiveTime
                        builtTgThread.add(
                            StationStopEntry(
                                stationName = stName,
                                telegramTime = effectiveTime,
                                isInitialStation = isInitial,
                                isFinalStation = isFinal,
                                isSkippedByTelegram = false,
                                isTechStop = false
                            )
                        )
                    } else if (isSingleStationMention) {
                        val isGlobalNoStop = Regex("(?iu)(без\\s+остановки|тех\\.?\\s*стоянк|без\\s+посадки)").containsMatchIn(cleanPara)
                        if (isGlobalNoStop || item.hasTechOrNoPassNote) {
                            if (!isClosedSt) skippedStops.add(stName)
                            if (item.hasTechOrNoPassNote) techStopStations.add(stName)
                        } else {
                            stationDepTimes[stName] = item.t2.ifEmpty { item.t1 }
                        }
                    } else {
                        val hasOnlyOneTime = item.t2.isEmpty()
                        if (hasOnlyOneTime || item.hasTechOrNoPassNote) {
                            if (!isClosedSt) skippedStops.add(stName)
                            if (item.hasTechOrNoPassNote) techStopStations.add(stName)
                            val labelTime = if (item.hasTechOrNoPassNote) {
                                val dep = item.techDepTime.ifEmpty { item.t2.ifEmpty { item.t1 } }
                                "$dep (тех. стоянка)"
                            } else {
                                "${item.t1} (без ост.)"
                            }
                            builtTgThread.add(
                                StationStopEntry(
                                    stationName = stName,
                                    telegramTime = labelTime,
                                    isInitialStation = false,
                                    isFinalStation = false,
                                    isSkippedByTelegram = true,
                                    isTechStop = item.hasTechOrNoPassNote
                                )
                            )
                        } else {
                            stationDepTimes[stName] = item.t2
                            builtTgThread.add(
                                StationStopEntry(
                                    stationName = stName,
                                    telegramTime = item.t2,
                                    isInitialStation = false,
                                    isFinalStation = false,
                                    isSkippedByTelegram = false,
                                    isTechStop = false
                                )
                            )
                        }
                    }
                }

                if (builtTgThread.isNotEmpty()) {
                    for (tNum in trainMatches) {
                        globalThreadsByTrain[tNum] = builtTgThread
                    }
                }

                val mentionedStations = PolygonConfig.ALL_STATIONS.filter { cleanPara.contains(it, true) }
                val targetStations = when {
                    rawStopItems.isNotEmpty() -> (rawStopItems.map { it.stationName } + mentionedStations).toSet()
                    mentionedStations.isNotEmpty() -> mentionedStations.toSet()
                    isFullCancel || isAppointed -> setOf("Крюково", "Химки", "Сходня", "Клин", "Подсолнечная", "Тверь")
                    else -> emptySet()
                }

                if (targetStations.isEmpty()) continue

                val validSkippedStops = skippedStops.distinct().filter { it !in PolygonConfig.HARD_CLOSED_STATIONS }
                val validAddedStops = addedStops.distinct().filter { it !in PolygonConfig.HARD_CLOSED_STATIONS }

                for (trainNum in trainMatches) {
                    val afishaTrain = afishaMap[trainNum]
                    val isToMsk = PolygonConfig.isTrainToMoscow(trainNum, cleanPara)
                    val routeStr = extractSmartRoute(cleanPara, trainNum, isToMsk, afishaTrain)

                    for (st in targetStations) {
                        if (st == "Левобережная" && !isToMsk) continue

                        val explicitStationBaseTime = if (rawStopItems.size <= 1) {
                            Regex("(?iu)отправлением\\s+в\\s+(\\d{1,2}[:.]\\d{2})").find(cleanPara)?.groupValues?.get(1)?.replace('.', ':')?.padStart(5, '0')
                        } else null
                        val afishaT = afishaTrain?.stationTimes?.get(st) ?: explicitStationBaseTime ?: "—"

                        val hasGeneralSkip = Regex("(?iu)проследует\\s+без\\s+остановки").containsMatchIn(cleanPara) &&
                            validSkippedStops.isEmpty() && mentionedStations.size == 1 && mentionedStations.first() == st
                        val isTechHere = st in techStopStations
                        val isSkippedHere = st in validSkippedStops || hasGeneralSkip || isTechHere
                        val isCanceledAtThisStation = isFullCancel || isSkippedHere

                        val rawItemForSt = rawStopItems.firstOrNull { it.stationName == st }
                        val rawTgT = stationDepTimes[st]
                            ?: rawItemForSt?.let { it.techDepTime.ifEmpty { it.t2.ifEmpty { it.t1 } } }
                            ?: afishaT

                        val tgT = when {
                            isFullCancel -> if (afishaT != "—") "$afishaT ОТМЕНА" else "ОТМЕНА"
                            isTechHere -> if (rawTgT != "—") "$rawTgT ТЕХ.СТ." else "ТЕХ. СТОЯНКА"
                            isSkippedHere -> if (rawTgT != "—") "$rawTgT БЕЗ ОСТ." else "БЕЗ ОСТАНОВКИ"
                            else -> rawTgT
                        }

                        val isMcd = PolygonConfig.isMcdTrain(trainNum, cleanPara)
                        val isExp = PolygonConfig.isExpressTrain(trainNum, cleanPara)

                        val finalSkippedForNote = validSkippedStops.filter { it != st && isFutureStop(st, it, isToMsk) }
                        val finalAddedForNote = validAddedStops.filter { it != st && isFutureStop(st, it, isToMsk) }

                        val remNoteUI = buildRemainingStopsNote(finalSkippedForNote, finalAddedForNote, isTechHere, isCanceledAtThisStation, pathStr, forWord = false)
                        val remNoteWord = buildRemainingStopsNote(finalSkippedForNote, finalAddedForNote, isTechHere, isCanceledAtThisStation, pathStr, forWord = true)

                        val changeSummary = when {
                            isFullCancel -> "Поезд №$trainNum ПОЛНОСТЬЮ ОТМЕНЁН (ст. $st)"
                            isTechHere -> "Поезд №$trainNum: тех. стоянка по ст. $st (без посадки/высадки пассажиров)"
                            isSkippedHere -> "Поезд №$trainNum проследует ст. $st БЕЗ ОСТАНОВКИ ($rawTgT)"
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
                            } else if (rawTgT != "—") {
                                "№$trainNum $routeStr отправлением в $rawTgT"
                            } else {
                                "№$trainNum $routeStr"
                            }
                            if (remNoteWord.isNotEmpty()) "$base. $remNoteWord".replace(" .", ".").trim() else base
                        }

                        if (isCanceledAtThisStation && !hasPartialRouteCancel) {
                            wordText = wordText.replace(Regex("(?iu)\\s*\\(?проследует\\s+без\\s+остановки\\)?"), "")
                            if (!Regex("(?iu)отмен").containsMatchIn(wordText)) {
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
                                hasExport = false,
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
                                telegramExcerpts = listOf(excerptItem),
                                rawExportSnippet = "",
                                telegramThread = builtTgThread,
                                auditStatus = AuditStatus.ONLY_TELEGRAM
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
                innerDateMap[dt] = mergeSameTrainRowsForDay(list, globalExcerptsByTrain, globalThreadsByTrain).associateBy { it.trainNumber }
            }
            finalMap[st] = innerDateMap
        }
        return TelegramParseResult(finalMap, globalExcerptsByTrain, globalThreadsByTrain)
    }

    fun parseSummaryExportFiles(
        files: List<File>,
        afishaMap: Map<String, AfishaBaseTrain>,
        telegramResult: TelegramParseResult
    ): StationDateTrainMap {
        if (files.isEmpty()) return telegramResult.stationMap

        val mutableResult = mutableMapOf<String, MutableMap<String, MutableList<TrainChangeRow>>>()
        for ((st, dtMap) in telegramResult.stationMap) {
            val mDt = mutableMapOf<String, MutableList<TrainChangeRow>>()
            for ((dt, trMap) in dtMap) mDt[dt] = trMap.values.toMutableList()
            mutableResult[st] = mDt
        }

        val trainNumberRegex = Regex("(?iu)(?:№|нр\\.?)\\s*([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)")

        for (file in files) {
            val text = readSmartTextFile(file)
            var currentStation = ""
            var currentDates = mutableListOf<String>()

            PolygonConfig.ALL_STATIONS.firstOrNull { file.nameWithoutExtension.contains(it, ignoreCase = true) }?.let {
                currentStation = it
            }

            fun detectStationHeaderLine(rawLine: String, prevLine: String): String? {
                val trimmed = rawLine.trim()
                if (trimmed.isEmpty()) return null
                if (trainNumberRegex.containsMatchIn(trimmed)) return null
                val prevTrim = prevLine.trim()
                if (prevTrim.endsWith(":") || prevTrim.endsWith(",")) return null
                if (trimmed.endsWith(".") && !trimmed.contains("Изменения", ignoreCase = true)) return null
                if (trimmed.contains(",")) return null

                val cleaned = trimmed
                    .replace(Regex("(?iu)Изменения\\s+в\\s+графике.*"), "")
                    .replace(Regex("(?iu)продолжение\\s+на.*"), "")
                    .replace(Regex("(?iu)(?:на\\s+Москву|от\\s+Москвы|из\\s+Москвы|на\\s+Ипподром|от\\s+Ипподрома).*"), "")
                    .replace(Regex("[():/\\-–—]+$"), "")
                    .trim()

                val norm = PolygonConfig.normalizeStationName(cleaned)
                return if (norm in PolygonConfig.ALL_STATIONS) norm else null
            }

            fun detectDateHeaderLine(rawLine: String): List<String>? {
                val trimmed = rawLine.trim()
                if (trimmed.isEmpty()) return null
                if (trainNumberRegex.containsMatchIn(trimmed)) return null
                if (trimmed.length > 55) return null
                val dates = extractSmartDates(trimmed)
                return if (dates.isNotEmpty()) dates else null
            }

            data class ExportBlock(
                val station: String,
                val dates: List<String>,
                val trainLine: String
            )

            val blocks = mutableListOf<ExportBlock>()
            var pendingTrainLine = ""
            var prevRawLine = ""

            fun flushPendingTrain() {
                if (pendingTrainLine.isNotBlank()) {
                    val st = currentStation.ifEmpty { "Крюково" }
                    val dts = if (currentDates.isNotEmpty()) currentDates.toList() else listOf("Без даты")
                    blocks.add(ExportBlock(st, dts, pendingTrainLine.trim()))
                    pendingTrainLine = ""
                }
            }

            for (line in text.lines()) {
                val trimmed = line.trim()
                if (trimmed.isEmpty()) continue

                if (Regex("(?iu)^(на\\s+Москву|от\\s+Москвы|на\\s+Область|Изменения\\s+в\\s+графике|продолжение\\s+на).*").matches(trimmed) &&
                    !trainNumberRegex.containsMatchIn(trimmed) &&
                    detectStationHeaderLine(trimmed, prevRawLine) == null
                ) {
                    prevRawLine = trimmed
                    continue
                }

                val stHeader = detectStationHeaderLine(trimmed, prevRawLine)
                if (stHeader != null) {
                    flushPendingTrain()
                    currentStation = stHeader
                    prevRawLine = trimmed
                    continue
                }

                val dtHeader = detectDateHeaderLine(trimmed)
                if (dtHeader != null) {
                    flushPendingTrain()
                    currentDates = dtHeader.toMutableList()
                    prevRawLine = trimmed
                    continue
                }

                if (trainNumberRegex.containsMatchIn(trimmed)) {
                    flushPendingTrain()
                    pendingTrainLine = trimmed
                } else {
                    if (pendingTrainLine.isNotEmpty()) {
                        pendingTrainLine += " $trimmed"
                    }
                }
                prevRawLine = trimmed
            }
            flushPendingTrain()

            for (block in blocks) {
                var cleanParagraph = block.trainLine
                    .replace(Regex("(?iu)[.\\s]*Изменений\\s*нет\\.?"), "")
                    .replace(Regex("(?iu)Бологое[- ]*Моск(?:овское)?\\.?"), "")
                    .replace(Regex("(?iu)\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)"), "")
                    .replace(Regex("\\.{2,}"), ".")
                    .replace(Regex("\\s{2,}"), " ")
                    .trim()

                if (cleanParagraph.isEmpty()) continue

                val hasDepartureWord = Regex("(?iu)(отправлением\\s+в|отправится\\s+в|назначается)").containsMatchIn(cleanParagraph)
                val hasArrivalWord = Regex("(?iu)(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)").containsMatchIn(cleanParagraph)
                if (hasArrivalWord && !hasDepartureWord) continue

                val effectiveStation = block.station
                val effectiveDates = block.dates

                val trNum = trainNumberRegex.find(cleanParagraph)?.groupValues?.get(1) ?: continue

                val hasPartialRouteCancel = Regex("(?iu)(отмен[а-я]*\\s+на\\s+участке|на\\s+участке\\s+[^.;]+\\s+отмен|следует\\s+до\\s+[А-Яа-яЁё\\-]+\\s+вместо|укорочен)").containsMatchIn(cleanParagraph)
                val localFullCancel = Regex("(?iu)(?:^|[^а-яА-ЯёЁ])(отменяется|отмен[её]н|отменить|полная\\s+отмена)(?:[^а-яА-ЯёЁ]|$)").containsMatchIn(cleanParagraph) &&
                    !Regex("(?iu)отменен[а-я]*\\s+остановк").containsMatchIn(cleanParagraph) &&
                    !Regex("(?iu)без\\s+остановк").containsMatchIn(cleanParagraph) &&
                    !hasPartialRouteCancel

                val isFullCancel = localFullCancel
                val isAppointed = Regex("(?iu)назначается|назначен\\s+поезд").containsMatchIn(cleanParagraph)

                cleanParagraph = cleanParagraph.replace(Regex("(?iu)\\s*\\(приб\\.?\\s*\\d{1,2}[:.]\\d{2}\\)"), "").trim()

                val skippedStops = mutableListOf<String>()
                skippedStops.addAll(extractStationsByRegex(cleanParagraph, "отменен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)"))
                skippedStops.addAll(extractStationsByRegex(cleanParagraph, "без\\s*останов[ок]+\\s*:?\\s*([^.;(]+)"))

                val addedStops = extractStationsByRegex(cleanParagraph, "назначен[ыа-я]*\\s*остановк[иа]?\\s*:?\\s*([^.;(]+)")

                val afishaTrain = afishaMap[trNum]
                val isToMsk = PolygonConfig.isTrainToMoscow(trNum, cleanParagraph)
                if (effectiveStation == "Левобережная" && !isToMsk) continue

                val routeStr = extractSmartRoute(cleanParagraph, trNum, isToMsk, afishaTrain)

                val hasGeneralSkip = Regex("(?iu)проследует\\s+без\\s+остановки").containsMatchIn(cleanParagraph) && skippedStops.isEmpty()
                val isSkippedHere = effectiveStation in skippedStops || hasGeneralSkip
                val isCanceledAtThisStation = isFullCancel || isSkippedHere

                val baseTimeMatch = Regex("(?iu)(?:отправлением\\s+в|назначается[^.\\n]*?\\s+в)\\s+(\\d{1,2}[:.]\\d{2})").find(cleanParagraph)
                val newTimeMatch = Regex("(?iu)(?:отправится|отпр\\.?|пойдет|следует)\\s+в\\s+(\\d{1,2}[:.]\\d{2})").find(cleanParagraph)
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

                var wordText = cleanParagraph.replace(Regex("(?iu)\\s*\\(?проследует\\s+без\\s+остановки\\)?"), "")
                if (isCanceledAtThisStation && !hasPartialRouteCancel && !Regex("(?iu)отмен").containsMatchIn(wordText)) {
                    wordText = "$wordText - ОТМЕНЕН"
                }
                wordText = wordText.replace(Regex("(?iu)\\s*(?:Малино|Петровско-Разумовская|Рижская|Лазурная)\\s*$"), "").trim()

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
                        hasExport = true,
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

        val out = mutableMapOf<String, Map<String, Map<String, TrainChangeRow>>>()
        for ((st, dtMap) in mutableResult) {
            val dOut = mutableMapOf<String, Map<String, TrainChangeRow>>()
            for ((dt, rows) in dtMap) {
                dOut[dt] = mergeSameTrainRowsForDay(
                    rows = rows,
                    globalExcerptsByTrain = telegramResult.excerptsByTrain,
                    globalThreadsByTrain = telegramResult.threadsByTrain
                ).associateBy { it.trainNumber }
            }
            out[st] = dOut
        }
        return out
    }

    fun parseYandexPdfFiles(files: List<File>): Pair<List<YandexParsedTrain>, String> {
        val parsed = mutableListOf<YandexParsedTrain>()
        val rawTextBuilder = StringBuilder()

        val trainNumRegex = Regex("\\b([1-7]\\d{2,3}(?:/[1-7]\\d{2,3})?)\\b")
        val timeRegex = Regex("\\b([0-2]?\\d:[0-5]\\d)\\b")
        val routeKeywordsRegex = Regex("(?iu)(Москва|Ленинградский|Тверь|Крюково|Зеленоград|Ипподром|47\\s*км|Клин|Подсолнечная|Сходня|Химки|Ховрино|Останкино|Конаково|Решетниково|Завидово|Поварово|МЦД|Ласточка|Экспресс|Стандарт|Комфорт|пригородный|везде|кроме|по\\s+расписанию|отмен[её]н|уш[её]л|изменен)")

        for (file in files) {
            val pagedChunks = readPagedTextChunks(file, sortPdfByPosition = true)
            val fullText = pagedChunks.joinToString("\n") { "--- Страница ${it.pageNumber} (${file.name}) ---\n${it.text}" }
            rawTextBuilder.appendLine(fullText)

            val detectedDate = extractSmartDates(fullText).firstOrNull()
                ?: extractSmartDates(file.name).firstOrNull()
                ?: "Все даты"

            val detectedStation = PolygonConfig.ALL_STATIONS.firstOrNull { st ->
                file.nameWithoutExtension.contains(st, ignoreCase = true) ||
                    fullText.lines().take(20).joinToString(" ").contains(st, ignoreCase = true)
            } ?: ""

            for (chunk in pagedChunks) {
                val lines = chunk.text.lines().map { it.trim() }.filter { it.isNotEmpty() }
                var i = 0
                while (i < lines.size) {
                    val line = lines[i]
                    if (line.startsWith("http", ignoreCase = true) || Regex("^\\d+/\\d+$").matches(line)) {
                        i++
                        continue
                    }

                    val timesInLine = timeRegex.findAll(line).map { it.groupValues[1].padStart(5, '0') }.toList()
                    val trainInLine = trainNumRegex.find(line)?.groupValues?.get(1)

                    if (timesInLine.isNotEmpty() && trainInLine != null) {
                        val depTime = timesInLine.first()
                        val arrTime = if (timesInLine.size >= 2) timesInLine[1] else "—"
                        val nextLineContext = lines.getOrNull(i + 1)?.takeIf { timeRegex.find(it) == null } ?: ""
                        val combinedInfo = "$line $nextLineContext".replace(Regex("\\s{2,}"), " ").trim()
                        val status = when {
                            Regex("(?iu)отмен").containsMatchIn(combinedInfo) -> "Отменён"
                            Regex("(?iu)изменен|опоздан|задерж").containsMatchIn(combinedInfo) -> "Изменено расписание"
                            else -> "По расписанию"
                        }
                        parsed.add(
                            YandexParsedTrain(
                                id = "${file.name}_${chunk.pageNumber}_${i}_${System.nanoTime()}",
                                trainNumber = trainInLine,
                                dateStr = detectedDate,
                                station = detectedStation,
                                departureTime = depTime,
                                arrivalTime = arrTime,
                                routeAndDetails = combinedInfo,
                                statusText = status,
                                sourceFileName = file.name,
                                pageNumber = chunk.pageNumber
                            )
                        )
                        i++
                        continue
                    }

                    if (timesInLine.isNotEmpty()) {
                        val nextLine = lines.getOrNull(i + 1) ?: ""
                        val prevLine = lines.getOrNull(i - 1) ?: ""
                        val trainInNext = if (timeRegex.find(nextLine) == null) trainNumRegex.find(nextLine)?.groupValues?.get(1) else null
                        val trainInPrev = if (timeRegex.find(prevLine) == null) trainNumRegex.find(prevLine)?.groupValues?.get(1) else null
                        val matchedTrain = trainInNext ?: trainInPrev

                        val combinedWindow = "$line $nextLine".replace(Regex("\\s{2,}"), " ").trim()
                        if (matchedTrain != null || routeKeywordsRegex.containsMatchIn(combinedWindow)) {
                            val depTime = timesInLine.first()
                            val arrTime = if (timesInLine.size >= 2) timesInLine[1] else "—"
                            val status = when {
                                Regex("(?iu)отмен").containsMatchIn(combinedWindow) -> "Отменён"
                                Regex("(?iu)изменен|опоздан|задерж").containsMatchIn(combinedWindow) -> "Изменено расписание"
                                else -> "По расписанию"
                            }
                            parsed.add(
                                YandexParsedTrain(
                                    id = "${file.name}_${chunk.pageNumber}_${i}_${System.nanoTime()}",
                                    trainNumber = matchedTrain ?: "—",
                                    dateStr = detectedDate,
                                    station = detectedStation,
                                    departureTime = depTime,
                                    arrivalTime = arrTime,
                                    routeAndDetails = combinedWindow,
                                    statusText = status,
                                    sourceFileName = file.name,
                                    pageNumber = chunk.pageNumber
                                )
                            )
                            if (trainInNext != null) i += 2 else i++
                            continue
                        }
                    }
                    i++
                }
            }
        }

        val deduplicated = parsed.distinctBy { "${it.sourceFileName}|${it.departureTime}|${it.trainNumber}|${it.routeAndDetails.take(40)}" }
            .sortedBy { parseTrainTimeForSorting(it.departureTime) }

        return Pair(deduplicated, rawTextBuilder.toString().trim())
    }
}

// ============================================================================
// 6. ГЕНЕРАТОР ФАЙЛОВ WORD (.DOCX) — МАКСИМАЛЬНО ПЛОТНОЕ ЗАПОЛНЕНИЕ БЕЗ РАЗРЫВОВ ВНИЗУ
// ============================================================================

object WordCancelGenerator {
    private val arrivalExcludeRegex = Regex("(?iu)(прибытием\\s+в|прибывает\\s+в|прибудет\\s+в)")
    private val departureIncludeRegex = Regex("(?iu)(отправлением\\s+в|отправится\\s+в|назначается)")
    private val nightNoteRegex = Regex("(?iu)\\s*\\(в\\s+ночь\\s+на\\s+[^)]+\\)")

    // Точная калибровка по фактической верстке MS Word (образцы Крюково, Сходня, Химки):
    // При полях сверху/снизу 450 twips (~0.8 см) физическая страница вмещает 53–54 строки,
    // а порог 52 визуальные строки исключает самопроизвольный перенос 1–2 строк в Word без заголовка
    // и появление полупустых страниц из-за двойного разрыва (рис. 1–2).
    private const val MAX_VISUAL_LINES_PER_PAGE = 52
    private const val CHARS_PER_LINE_12PT = 90

    private fun isTerminalStationForTrain(station: String, row: TrainChangeRow): Boolean {
        val text = row.wordExportLine
        if (departureIncludeRegex.containsMatchIn(text)) return false
        if (arrivalExcludeRegex.containsMatchIn(text)) return true

        val destPattern = Regex("(?iu)^\\s*№\\s*[\\d/]+\\s+(?:из\\s+[А-Яа-яЁё\\-]+\\s+)?до\\s+([А-Яа-яЁё\\-]+)")
        val match = destPattern.find(text)

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

    private fun hasAnyExportRows(allData: StationDateTrainMap): Boolean {
        return allData.values.any { dtMap -> dtMap.values.any { trMap -> trMap.values.any { it.hasExport } } }
    }

    /**
     * Устанавливает компактные поля страницы Word (верхнее и нижнее по 1.0 см = 567 twips),
     * чтобы текст доходил почти до самого низа листа и не оставлял пустых мест (рис. 1-3).
     */
    private fun configureCompactPageMargins(doc: XWPFDocument) {
        try {
            val body = doc.document.body
            val sectPr = if (body.isSetSectPr) body.sectPr else body.addNewSectPr()
            val pageMar = if (sectPr.isSetPgMar) sectPr.pgMar else sectPr.addNewPgMar()
            pageMar.top = BigInteger.valueOf(450)    // ~0.8 см сверху (резерв от авто-переноса Word)
            pageMar.bottom = BigInteger.valueOf(450) // ~0.8 см снизу (резерв от авто-переноса Word)
            pageMar.left = BigInteger.valueOf(850)   // 1.5 см слева
            pageMar.right = BigInteger.valueOf(567)  // 1.0 см справа
        } catch (_: Exception) {}
    }

    private fun estimateVisualLines(textWithIndent: String): Int {
        return maxOf(1, (textWithIndent.length + CHARS_PER_LINE_12PT - 1) / CHARS_PER_LINE_12PT)
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

        val onlyExportRows = hasAnyExportRows(allData)
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
                val filteredRows = trainMap.values.filter { row ->
                    val sourceAllowed = if (onlyExportRows) row.hasExport else true
                    val isNotArrivalOnly = departureIncludeRegex.containsMatchIn(row.wordExportLine) || !arrivalExcludeRegex.containsMatchIn(row.wordExportLine)
                    sourceAllowed &&
                        row.isToMoscow == isToMoscow &&
                        isNotArrivalOnly &&
                        !isTerminalStationForTrain(station, row)
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
                configureCompactPageMargins(doc)
                val subtitle = getDirectionSubtitle(isToMoscow)
                addCenteredBoldHeader(doc, "$station  Изменения в графике движения эл/поездов")
                addCenteredBoldHeader(doc, subtitle)
                var currentPageVisualLines = 2

                for ((_, groupedDates) in signatureToDates) {
                    val formattedDatesHeader = formatMergedDatesHeader(groupedDates)
                    val linesForDate = dateToCleanLines[groupedDates.first()] ?: emptyList()
                    if (linesForDate.isEmpty()) continue

                    val firstTrainVisualLines = estimateVisualLines("        ${linesForDate.first().trim()}")
                    val secondTrainVisualLines = if (linesForDate.size > 1) {
                        estimateVisualLines("        ${linesForDate[1].trim()}")
                    } else 0

                    // Если в самом низу страницы помещается только заголовок новой даты и 1-2 строчки,
                    // после чего сразу потребуется разрыв страницы — начинаем новую дату с новой страницы.
                    if (currentPageVisualLines + 1 + firstTrainVisualLines + secondTrainVisualLines > MAX_VISUAL_LINES_PER_PAGE) {
                        addNewPageStationHeader(doc, station, subtitle)
                        currentPageVisualLines = 2
                    }

                    addCenteredBoldHeader(doc, formattedDatesHeader)
                    currentPageVisualLines += 1

                    for ((idx, rawLine) in linesForDate.withIndex()) {
                        val finalTrainLine = "        ${rawLine.trim()}"
                        val neededLines = estimateVisualLines(finalTrainLine)
                        val isLastTrainOfDate = idx == linesForDate.lastIndex

                        // Допускаем +1 резервную строку для самого последнего поезда даты,
                        // чтобы не переносить 1 одиночную строчку на новую страницу
                        val effectiveLimit = if (isLastTrainOfDate && neededLines == 1) {
                            MAX_VISUAL_LINES_PER_PAGE + 1
                        } else {
                            MAX_VISUAL_LINES_PER_PAGE
                        }

                        if (currentPageVisualLines + neededLines > effectiveLimit) {
                            addContinuationHeader(doc, station, formattedDatesHeader, subtitle)
                            currentPageVisualLines = 2
                        }

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
                        run.setText(finalTrainLine)
                        currentPageVisualLines += neededLines
                    }
                }
                FileOutputStream(outFile).use { doc.write(it) }
                generatedCount++
            }
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

    private fun addNewPageStationHeader(doc: XWPFDocument, station: String, subtitle: String) {
        addCenteredBoldHeader(doc, "$station  Изменения в графике движения эл/поездов", pageBreakBefore = true)
        addCenteredBoldHeader(doc, subtitle, pageBreakBefore = false)
    }

    /**
     * Двухстрочный заголовок при продолжении списка поездов одной даты на следующей странице:
     * 1-я строка: "$station продолжение на $datesHeader"
     * 2-я строка: "$subtitle"
     */
    private fun addContinuationHeader(doc: XWPFDocument, station: String, datesHeader: String, subtitle: String) {
        addCenteredBoldHeader(doc, "$station продолжение на $datesHeader", pageBreakBefore = true)
        addCenteredBoldHeader(doc, subtitle, pageBreakBefore = false)
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
fun TrainRApp() {
    val scope = rememberCoroutineScope()

    var currentNavTab: MainNavTab by remember { mutableStateOf(MainNavTab.UPLOADS) }
    var currentDirectionTab: DirectionTab by remember { mutableStateOf(DirectionTab.TO_MOSCOW) }
    var currentTrainFilter: TrainFilterType by remember { mutableStateOf(TrainFilterType.ALL) }
    var currentReportFilter: ReportFilterType by remember { mutableStateOf(ReportFilterType.ALL) }

    var searchQuery: String by remember { mutableStateOf("") }
    var systemStatusText: String by remember { mutableStateOf("Система готова. Загрузите файлы выгрузки (.txt/.docx) для компиляции в Word.") }

    var afishaCacheMap: Map<String, AfishaBaseTrain> by remember { mutableStateOf(emptyMap()) }
    var stationScheduleMap: StationDateTrainMap by remember { mutableStateOf(emptyMap()) }
    var yandexTrainsList: List<YandexParsedTrain> by remember { mutableStateOf(emptyList()) }
    var yandexRawPdfText: String by remember { mutableStateOf("") }

    var loadedTelegramFiles by remember { mutableStateOf(listOf<File>()) }
    var loadedExportFiles by remember { mutableStateOf(listOf<File>()) }
    var loadedYandexFiles by remember { mutableStateOf(listOf<File>()) }

    var showFileManagerDialog by remember { mutableStateOf(false) }

    var activePreviewStation: String by remember { mutableStateOf("Крюково") }
    var checkedStationsForWord: Set<String> by remember { mutableStateOf(setOf("Крюково", "Химки", "Сходня")) }
    var selectedFilterDates: Set<String> by remember { mutableStateOf(emptySet()) }

    var showCalendarDialog: Boolean by remember { mutableStateOf(false) }
    var selectedTrainForDossier: TrainChangeRow? by remember { mutableStateOf(null) }

    fun refreshData() {
        scope.launch(Dispatchers.IO) {
            val tgResult = if (loadedTelegramFiles.isNotEmpty()) {
                FileParsers.parseTelegramFiles(loadedTelegramFiles, afishaCacheMap)
            } else {
                TelegramParseResult(emptyMap(), emptyMap(), emptyMap())
            }
            val combinedMap = if (loadedExportFiles.isNotEmpty()) {
                FileParsers.parseSummaryExportFiles(loadedExportFiles, afishaCacheMap, tgResult)
            } else {
                tgResult.stationMap
            }
            val (parsedYandex, rawYandexText) = if (loadedYandexFiles.isNotEmpty()) {
                FileParsers.parseYandexPdfFiles(loadedYandexFiles)
            } else {
                Pair(emptyList(), "")
            }
            val built = buildFullThreads(combinedMap, afishaCacheMap, parsedYandex)
            withContext(Dispatchers.Main) {
                yandexTrainsList = parsedYandex
                yandexRawPdfText = rawYandexText
                stationScheduleMap = built
                val firstStationWithData = PolygonConfig.ALL_STATIONS.firstOrNull { built[it]?.isNotEmpty() == true }
                if (built[activePreviewStation].isNullOrEmpty() && firstStationWithData != null) {
                    activePreviewStation = firstStationWithData
                }
                val tgInfo = if (loadedTelegramFiles.isEmpty()) "ТГ: не загружена" else "ТГ: ${loadedTelegramFiles.size} шт."
                val expInfo = if (loadedExportFiles.isEmpty()) "Выгрузки: 0" else "Выгрузки: ${loadedExportFiles.size} шт."
                val yndInfo = if (loadedYandexFiles.isEmpty()) "" else " | Яндекс PDF: ${parsedYandex.size} зап."
                systemStatusText = "Данные скомпилированы ($expInfo | $tgInfo$yndInfo)"
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
                        systemStatusText = "Загружен кэш базовой Афиши (${parsed.size} поездов)."
                    }
                }
            }
        }
    }

    val allDatesWithChanges: Set<String> = remember(stationScheduleMap, loadedExportFiles) {
        val hasExp = loadedExportFiles.isNotEmpty()
        stationScheduleMap.values.flatMap { dtMap ->
            dtMap.entries.filter { (_, trMap) ->
                if (hasExp) trMap.values.any { it.hasExport } else trMap.isNotEmpty()
            }.map { it.key }
        }.toSet()
    }

    val baseRows: List<TrainChangeRow> = remember(
        stationScheduleMap,
        activePreviewStation,
        selectedFilterDates,
        currentDirectionTab,
        currentTrainFilter,
        searchQuery,
        loadedExportFiles,
        currentNavTab
    ) {
        val dateMap = stationScheduleMap[activePreviewStation] ?: emptyMap()
        val datesToInclude = if (selectedFilterDates.isEmpty()) dateMap.keys else dateMap.keys.intersect(selectedFilterDates)
        val hasExp = loadedExportFiles.isNotEmpty()

        datesToInclude.sortedBy { parseDateForSort(it) }.flatMap { dt -> dateMap[dt]?.values ?: emptyList() }
            .filter { row ->
                if (currentNavTab == MainNavTab.UPLOADS && hasExp) row.hasExport else true
            }
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
                    contentPadding = PaddingValues(vertical = 12.dp, horizontal = 8.dp)
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
                        yandexRawPdfText = ""
                        loadedTelegramFiles = emptyList()
                        loadedExportFiles = emptyList()
                        loadedYandexFiles = emptyList()
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
            Surface(color = Color.White, shadowElevation = 2.dp, modifier = Modifier.fillMaxWidth()) {
                Row(
                    modifier = Modifier.fillMaxWidth().padding(horizontal = 16.dp, vertical = 10.dp),
                    verticalAlignment = Alignment.CenterVertically,
                    horizontalArrangement = Arrangement.SpaceBetween
                ) {
                    SearchInputBox(
                        value = searchQuery,
                        onValueChange = { searchQuery = it },
                        placeholder = "Поиск № поезда, маршрута...",
                        modifier = Modifier.width(290.dp).height(42.dp)
                    )

                    Row(horizontalArrangement = Arrangement.spacedBy(10.dp), verticalAlignment = Alignment.CenterVertically) {
                        OutlinedButton(
                            onClick = { showFileManagerDialog = true },
                            shape = RoundedCornerShape(8.dp),
                            border = BorderStroke(1.dp, TrainRTheme.NavyDark),
                            contentPadding = PaddingValues(horizontal = 12.dp, vertical = 8.dp)
                        ) {
                            Icon(Icons.Default.Settings, contentDescription = null, tint = TrainRTheme.NavyDark, modifier = Modifier.size(16.dp))
                            Spacer(modifier = Modifier.width(6.dp))
                            Text("ФАЙЛЫ (${loadedTelegramFiles.size + loadedExportFiles.size + loadedYandexFiles.size})", color = TrainRTheme.NavyDark, fontWeight = FontWeight.Bold, fontSize = 12.sp)
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

                        TopBarLoadButton("ВЫГРУЗКИ (.TXT)", loadedExportFiles.size, TrainRTheme.GreenSuccess) {
                            val files = pickFilesDialog("Выберите файлы выгрузок .txt / .docx (можно несколько одновременно)", arrayOf("txt", "docx", "pdf"), true)
                            if (files.isNotEmpty()) {
                                loadedExportFiles = (loadedExportFiles + files).distinctBy { it.absolutePath }
                                systemStatusText = "Добавлены выгрузки (${loadedExportFiles.size} шт.). Компиляция по дням и станциям..."
                                refreshData()
                            }
                        }
                    }
                }
            }

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
                                    systemStatusText = "Выполнена сверка источников по ст. $activePreviewStation."
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
                            "Станции полигона (чекбокс — для выгрузки в Word, клик по названию — активная станция отчета):",
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
                        yandexRawText = yandexRawPdfText,
                        loadedYandexFiles = loadedYandexFiles,
                        searchQuery = searchQuery,
                        onLoadYandexPdf = {
                            val files = pickFilesDialog("Выберите PDF с расписанием Яндекса", arrayOf("pdf", "txt"), true)
                            if (files.isNotEmpty()) {
                                loadedYandexFiles = (loadedYandexFiles + files).distinctBy { it.absolutePath }
                                systemStatusText = "Загружен PDF Яндекс.Расписаний (${loadedYandexFiles.size} шт.). Чтение данных..."
                                refreshData()
                            }
                        },
                        onRunYandexAudit = {
                            if (loadedYandexFiles.isNotEmpty()) {
                                refreshData()
                                systemStatusText = "Сверка с PDF Яндекс.Расписаний по ст. $activePreviewStation выполнена (${yandexTrainsList.size} записей)."
                            }
                        },
                        onClearYandexPdf = {
                            loadedYandexFiles = emptyList()
                            yandexTrainsList = emptyList()
                            yandexRawPdfText = ""
                            refreshData()
                            systemStatusText = "Файлы PDF Яндекса очищены."
                        },
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
            yandexFiles = loadedYandexFiles,
            onRemoveTelegram = { file ->
                loadedTelegramFiles = loadedTelegramFiles - file
                refreshData()
            },
            onRemoveExport = { file ->
                loadedExportFiles = loadedExportFiles - file
                refreshData()
            },
            onRemoveYandex = { file ->
                loadedYandexFiles = loadedYandexFiles - file
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
fun SearchInputBox(
    value: String,
    onValueChange: (String) -> Unit,
    placeholder: String,
    modifier: Modifier = Modifier
) {
    Surface(
        shape = RoundedCornerShape(8.dp),
        color = Color.White,
        border = BorderStroke(1.dp, TrainRTheme.BorderCol),
        modifier = modifier
    ) {
        Row(
            verticalAlignment = Alignment.CenterVertically,
            modifier = Modifier.fillMaxSize().padding(horizontal = 12.dp)
        ) {
            Icon(
                imageVector = Icons.Default.Search,
                contentDescription = "Поиск",
                tint = TrainRTheme.TextSub,
                modifier = Modifier.size(18.dp)
            )
            Spacer(modifier = Modifier.width(8.dp))
            Box(
                contentAlignment = Alignment.CenterStart,
                modifier = Modifier.weight(1f).fillMaxHeight()
            ) {
                if (value.isEmpty()) {
                    Text(
                        text = placeholder,
                        color = TrainRTheme.TextSub,
                        fontSize = 13.sp,
                        maxLines = 1,
                        overflow = TextOverflow.Ellipsis
                    )
                }
                BasicTextField(
                    value = value,
                    onValueChange = onValueChange,
                    singleLine = true,
                    textStyle = TextStyle(
                        color = TrainRTheme.TextDark,
                        fontSize = 13.sp,
                        fontWeight = FontWeight.Medium
                    ),
                    cursorBrush = SolidColor(TrainRTheme.RoyalBlue),
                    modifier = Modifier.fillMaxWidth()
                )
            }
            if (value.isNotEmpty()) {
                Icon(
                    imageVector = Icons.Default.Clear,
                    contentDescription = "Очистить",
                    tint = TrainRTheme.TextSub,
                    modifier = Modifier
                        .size(18.dp)
                        .clickable { onValueChange("") }
                )
            }
        }
    }
}

@Composable
fun FileManagerDialog(
    telegrams: List<File>,
    exports: List<File>,
    yandexFiles: List<File>,
    onRemoveTelegram: (File) -> Unit,
    onRemoveExport: (File) -> Unit,
    onRemoveYandex: (File) -> Unit,
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
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, contentDescription = "Закрыть") }
                }
                Spacer(modifier = Modifier.height(10.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)

                Column(modifier = Modifier.weight(1f, fill = false).verticalScroll(rememberScrollState()).padding(vertical = 12.dp)) {
                    Text("Файлы выгрузок .txt / .docx (${exports.size})", fontWeight = FontWeight.Bold, color = TrainRTheme.GreenSuccess, fontSize = 13.sp)
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

                    Spacer(modifier = Modifier.height(14.dp))

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

                    Spacer(modifier = Modifier.height(14.dp))

                    Text("PDF Яндекс.Расписаний (${yandexFiles.size})", fontWeight = FontWeight.Bold, color = TrainRTheme.YandexRed, fontSize = 13.sp)
                    if (yandexFiles.isEmpty()) Text("Файлы PDF Яндекса не загружены", fontSize = 12.sp, color = TrainRTheme.TextSub)
                    yandexFiles.forEach { file ->
                        Row(
                            modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp),
                            horizontalArrangement = Arrangement.SpaceBetween,
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text(file.name, fontSize = 12.sp, maxLines = 1, overflow = TextOverflow.Ellipsis, modifier = Modifier.weight(1f))
                            IconButton(onClick = { onRemoveYandex(file) }, modifier = Modifier.size(24.dp)) {
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
        contentPadding = PaddingValues(horizontal = 12.dp, vertical = 8.dp)
    ) {
        Icon(imageVector = Icons.Default.Add, contentDescription = label, tint = TrainRTheme.TextDark, modifier = Modifier.size(15.dp))
        Spacer(modifier = Modifier.width(5.dp))
        Text(text = label, color = TrainRTheme.TextDark, fontWeight = FontWeight.Bold, fontSize = 11.sp)
        Spacer(modifier = Modifier.width(6.dp))
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
                        "ℹ Телеграмма еще не загружена. Отображаются скомпилированные данные из загруженных выгрузок по ст. $activeStation.",
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
                Text("СТРОКА ВЫГРУЗКИ (ДЛЯ WORD)", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(2.3f))
                Text("СТАТУС", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
            }
            if (rows.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    Text("Нет данных об изменениях по станции «$activeStation». Загрузите файлы выгрузки (.txt/.docx) или телеграмму.", color = TrainRTheme.TextSub, fontSize = 14.sp)
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
                            val isTgCanceled = row.hasTelegram && (row.isCanceled || row.telegramTime.contains("ОТМЕН", true) || row.telegramTime.contains("БЕЗ ОСТ", true) || row.telegramTime.contains("ТЕХ.", true))
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
                                Text(row.wordExportLine.ifBlank { row.actualChangesText }, fontSize = 12.sp, color = TrainRTheme.TextDark)
                                if (row.telegramExcerpts.isNotEmpty()) {
                                    Text(
                                        "Найдено в ТГ (${row.telegramExcerpts.size}): ${row.telegramExcerpts.joinToString("; ") { "${it.telegramNumber} (${it.telegramPage})" }}",
                                        fontSize = 10.sp,
                                        color = TrainRTheme.AmberWarn,
                                        fontWeight = FontWeight.Bold
                                    )
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
                            "ОТЧЕТ СФОРМИРОВАН ПО СТАНЦИИ: ${activeStation.uppercase()} • ФИЛЬТР: ${currentFilter.title} (Найдено: ${reportRows.size})",
                            fontSize = 14.sp,
                            fontWeight = FontWeight.ExtraBold,
                            color = TrainRTheme.NavyDark
                        )
                        Text(
                            "Нажмите на поезд, чтобы открыть полное досье и нитку следования по станции «$activeStation»",
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
                    Surface(color = TrainRTheme.NavyDark, shape = RoundedCornerShape(4.dp)) {
                        Text(
                            "СТ. ${row.station.uppercase()}",
                            color = Color.White,
                            fontSize = 10.sp,
                            fontWeight = FontWeight.ExtraBold,
                            modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp)
                        )
                    }
                    Spacer(modifier = Modifier.width(8.dp))
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
                    Text("Дата: ${row.dateStr}", fontSize = 12.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.SemiBold)
                }
                Surface(color = row.auditStatus.color, shape = RoundedCornerShape(6.dp)) {
                    Text(row.auditStatus.label, color = Color.White, fontSize = 10.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp))
                }
            }
            Spacer(modifier = Modifier.height(6.dp))
            val tgTimeDisplay = if (row.hasTelegram) row.telegramTime else "— (не загружена)"
            Text(
                "Сверка по ст. ${row.station}: Афиша ${row.afishaTime} ➔ Телеграмма $tgTimeDisplay ➔ Выгрузка ${row.exportTime} ➔ Яндекс ${row.yandexTime}",
                fontSize = 12.sp,
                fontFamily = FontFamily.Monospace,
                fontWeight = FontWeight.SemiBold,
                color = TrainRTheme.RoyalBlue
            )
            Spacer(modifier = Modifier.height(4.dp))
            Text("Строка для Word: ${row.wordExportLine}", fontSize = 12.sp, color = TrainRTheme.TextDark, fontWeight = FontWeight.Medium)

            Spacer(modifier = Modifier.height(8.dp))
            if (row.telegramExcerpts.isNotEmpty()) {
                Column(verticalArrangement = Arrangement.spacedBy(6.dp)) {
                    Text(
                        "Найдено упоминаний поезда №${row.trainNumber} в загруженных телеграммах: ${row.telegramExcerpts.size} шт.",
                        fontSize = 11.sp,
                        fontWeight = FontWeight.ExtraBold,
                        color = TrainRTheme.AmberWarn
                    )
                    row.telegramExcerpts.forEachIndexed { idx, item ->
                        Surface(
                            color = Color(0xFFFFFBEB),
                            shape = RoundedCornerShape(6.dp),
                            border = BorderStroke(1.dp, TrainRTheme.GoldAccent.copy(alpha = 0.6f)),
                            modifier = Modifier.fillMaxWidth()
                        ) {
                            Column(modifier = Modifier.padding(10.dp)) {
                                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(14.dp)) {
                                    Text(
                                        "ТГ #${idx + 1}: ${item.telegramNumber}",
                                        fontWeight = FontWeight.ExtraBold,
                                        fontSize = 12.sp,
                                        color = TrainRTheme.NavyDark
                                    )
                                    Text(
                                        "Дата ТГ: ${item.telegramDate}",
                                        fontWeight = FontWeight.Bold,
                                        fontSize = 12.sp,
                                        color = TrainRTheme.NavyDark
                                    )
                                    Text(
                                        "Страница: ${item.telegramPage}",
                                        fontWeight = FontWeight.ExtraBold,
                                        fontSize = 12.sp,
                                        color = TrainRTheme.RoyalBlue
                                    )
                                    Text(
                                        "Файл: ${item.fileName}",
                                        fontSize = 11.sp,
                                        color = TrainRTheme.TextSub
                                    )
                                }
                                Spacer(modifier = Modifier.height(4.dp))
                                Text(
                                    item.snippet,
                                    fontSize = 11.sp,
                                    color = TrainRTheme.TextDark,
                                    fontFamily = FontFamily.Monospace
                                )
                            }
                        }
                    }
                }
            } else {
                Text(
                    "Телеграмма: не загружена или данный поезд в ней не упоминается (источник: ${row.exportBasis})",
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
    yandexRawText: String,
    loadedYandexFiles: List<File>,
    searchQuery: String,
    onLoadYandexPdf: () -> Unit,
    onRunYandexAudit: () -> Unit,
    onClearYandexPdf: () -> Unit,
    onRowClick: (TrainChangeRow) -> Unit
) {
    var viewMode by remember { mutableStateOf(YandexViewMode.PDF_DATA) }

    val filteredYandexList = remember(yandexList, searchQuery) {
        if (searchQuery.isBlank()) yandexList
        else yandexList.filter {
            it.trainNumber.contains(searchQuery.trim(), ignoreCase = true) ||
                it.departureTime.contains(searchQuery.trim(), ignoreCase = true) ||
                it.routeAndDetails.contains(searchQuery.trim(), ignoreCase = true)
        }
    }

    val rowsByTrainNum = remember(rows) { rows.associateBy { it.trainNumber } }
    val rowsByBaseTime = remember(rows) { rows.associateBy { it.afishaTime } }

    Surface(color = Color.White, shape = RoundedCornerShape(10.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol), modifier = Modifier.fillMaxSize()) {
        Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                Column(modifier = Modifier.weight(1f)) {
                    Text(
                        "АУДИТ PDF-РАСПИСАНИЯ ЯНДЕКСА • СТАНЦИЯ: ${activeStation.uppercase()}",
                        fontSize = 15.sp,
                        fontWeight = FontWeight.ExtraBold,
                        color = TrainRTheme.NavyDark
                    )
                    Text(
                        "Загружено файлов PDF Яндекса: ${loadedYandexFiles.size} • Найдено рейсов в PDF: ${yandexList.size}",
                        fontSize = 12.sp,
                        color = TrainRTheme.TextSub
                    )
                }
                Row(horizontalArrangement = Arrangement.spacedBy(10.dp), verticalAlignment = Alignment.CenterVertically) {
                    Button(
                        onClick = onLoadYandexPdf,
                        colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.YandexRed),
                        shape = RoundedCornerShape(8.dp)
                    ) {
                        Icon(Icons.Default.Add, contentDescription = null, tint = Color.White, modifier = Modifier.size(16.dp))
                        Spacer(modifier = Modifier.width(6.dp))
                        Text("ЗАГРУЗИТЬ PDF ЯНДЕКС.РАСПИСАНИЙ", color = Color.White, fontWeight = FontWeight.ExtraBold, fontSize = 12.sp)
                    }

                    Button(
                        onClick = onRunYandexAudit,
                        enabled = loadedYandexFiles.isNotEmpty(),
                        colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue),
                        shape = RoundedCornerShape(8.dp)
                    ) {
                        Icon(Icons.Default.Refresh, contentDescription = null, tint = Color.White, modifier = Modifier.size(16.dp))
                        Spacer(modifier = Modifier.width(6.dp))
                        Text("СВЕРИТЬ", color = Color.White, fontWeight = FontWeight.ExtraBold, fontSize = 12.sp)
                    }

                    if (loadedYandexFiles.isNotEmpty()) {
                        OutlinedButton(
                            onClick = onClearYandexPdf,
                            shape = RoundedCornerShape(8.dp),
                            border = BorderStroke(1.dp, TrainRTheme.RedCancel)
                        ) {
                            Text("Очистить PDF", color = TrainRTheme.RedCancel, fontWeight = FontWeight.Bold, fontSize = 11.sp)
                        }
                    }
                }
            }

            Spacer(modifier = Modifier.height(10.dp))
            HorizontalDivider(color = TrainRTheme.BorderCol)
            Spacer(modifier = Modifier.height(10.dp))

            if (loadedYandexFiles.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    Column(horizontalAlignment = Alignment.CenterHorizontally) {
                        Text(
                            "Файл PDF с Яндекс.Расписанием еще не загружен — сверять нечего.",
                            fontSize = 15.sp,
                            fontWeight = FontWeight.ExtraBold,
                            color = TrainRTheme.NavyDark
                        )
                        Spacer(modifier = Modifier.height(6.dp))
                        Text(
                            "Нажмите кнопку «ЗАГРУЗИТЬ PDF ЯНДЕКС.РАСПИСАНИЙ» в правом верхнем углу этой вкладки для отображения данных PDF и сверки.",
                            fontSize = 13.sp,
                            color = TrainRTheme.TextSub
                        )
                    }
                }
            } else {
                Row(
                    modifier = Modifier.fillMaxWidth().padding(bottom = 10.dp),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        YandexViewMode.entries.forEach { mode ->
                            val selected = viewMode == mode
                            Button(
                                onClick = { viewMode = mode },
                                colors = ButtonDefaults.buttonColors(
                                    containerColor = if (selected) TrainRTheme.NavyDark else Color(0xFFF1F5F9),
                                    contentColor = if (selected) Color.White else TrainRTheme.TextDark
                                ),
                                shape = RoundedCornerShape(8.dp),
                                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp)
                            ) {
                                val countSuffix = when (mode) {
                                    YandexViewMode.PDF_DATA -> " (${filteredYandexList.size})"
                                    YandexViewMode.COMPARE_CHANGES -> " (${rows.size})"
                                    YandexViewMode.RAW_TEXT -> ""
                                }
                                Text("${mode.label}$countSuffix", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }
                        }
                    }
                    Text(
                        "Файлы: ${loadedYandexFiles.joinToString(", ") { it.name }}",
                        fontSize = 11.sp,
                        color = TrainRTheme.TextSub,
                        maxLines = 1,
                        overflow = TextOverflow.Ellipsis
                    )
                }

                when (viewMode) {
                    YandexViewMode.PDF_DATA -> {
                        if (filteredYandexList.isNotEmpty()) {
                            Column(modifier = Modifier.fillMaxSize().border(1.dp, TrainRTheme.BorderCol, RoundedCornerShape(8.dp))) {
                                Row(
                                    modifier = Modifier.fillMaxWidth().background(TrainRTheme.NavyDark).padding(horizontal = 12.dp, vertical = 9.dp),
                                    verticalAlignment = Alignment.CenterVertically
                                ) {
                                    Text("ОТПР. (PDF)", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(0.9f))
                                    Text("№ ПОЕЗДА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.0f))
                                    Text("ДАННЫЕ ИЗ PDF ЯНДЕКС.РАСПИСАНИЙ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(2.8f))
                                    Text("СТАТУС В PDF", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.1f))
                                    Text("ВЫГРУЗКА / ТГ", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.2f))
                                    Text("СВЕРКА", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.weight(1.3f))
                                }
                                LazyColumn(modifier = Modifier.fillMaxSize()) {
                                    items(filteredYandexList, key = { it.id }) { yTrain ->
                                        val matchedRow = if (yTrain.trainNumber != "—") {
                                            rowsByTrainNum[yTrain.trainNumber]
                                        } else {
                                            rowsByBaseTime[yTrain.departureTime]
                                        }
                                        val displayTrainNum = if (yTrain.trainNumber != "—") "№${yTrain.trainNumber}" else matchedRow?.let { "№${it.trainNumber}" } ?: "—"
                                        val ourStatusTime = matchedRow?.let {
                                            if (it.hasExport) it.exportTime else it.telegramTime
                                        } ?: "По графику"

                                        val isYandexCanceled = yTrain.statusText.contains("Отмен", true)
                                        val isOurCanceled = matchedRow?.isCanceled == true

                                        val (reconcileLabel, reconcileColor) = when {
                                            matchedRow == null -> Pair("БЕЗ ИЗМЕНЕНИЙ", TrainRTheme.TextSub)
                                            isOurCanceled && !isYandexCanceled -> Pair("ОШИБКА: НЕ ОТМЕНЕН В ЯНДЕКСЕ", TrainRTheme.RedCancel)
                                            !isOurCanceled && isYandexCanceled -> Pair("ОТМЕНЕН В ЯНДЕКСЕ", TrainRTheme.AmberWarn)
                                            matchedRow.exportTime != "—" && !isOurCanceled && !matchedRow.exportTime.contains(yTrain.departureTime) ->
                                                Pair("РАСХОЖДЕНИЕ (${matchedRow.exportTime})", TrainRTheme.AmberWarn)
                                            else -> Pair("СОВПАДАЕТ", TrainRTheme.GreenSuccess)
                                        }

                                        Row(
                                            modifier = Modifier
                                                .fillMaxWidth()
                                                .background(
                                                    when {
                                                        reconcileColor == TrainRTheme.RedCancel -> TrainRTheme.RedCancel.copy(alpha = 0.08f)
                                                        isYandexCanceled -> TrainRTheme.AmberWarn.copy(alpha = 0.06f)
                                                        else -> Color.White
                                                    }
                                                )
                                                .clickable { if (matchedRow != null) onRowClick(matchedRow) }
                                                .padding(horizontal = 12.dp, vertical = 8.dp),
                                            verticalAlignment = Alignment.CenterVertically
                                        ) {
                                            Text(
                                                yTrain.departureTime,
                                                fontSize = 13.sp,
                                                fontFamily = FontFamily.Monospace,
                                                fontWeight = FontWeight.ExtraBold,
                                                color = if (isYandexCanceled) TrainRTheme.RedCancel else TrainRTheme.NavyDark,
                                                modifier = Modifier.weight(0.9f)
                                            )
                                            Text(
                                                displayTrainNum,
                                                fontSize = 12.sp,
                                                fontWeight = FontWeight.ExtraBold,
                                                color = TrainRTheme.RoyalBlue,
                                                modifier = Modifier.weight(1.0f)
                                            )
                                            Column(modifier = Modifier.weight(2.8f).padding(end = 8.dp)) {
                                                Text(yTrain.routeAndDetails, fontSize = 12.sp, color = TrainRTheme.TextDark)
                                                Text("Стр. ${yTrain.pageNumber} • ${yTrain.sourceFileName}", fontSize = 10.sp, color = TrainRTheme.TextSub)
                                            }
                                            Text(
                                                yTrain.statusText,
                                                fontSize = 11.sp,
                                                fontWeight = FontWeight.Bold,
                                                color = if (isYandexCanceled) TrainRTheme.RedCancel else TrainRTheme.GreenSuccess,
                                                modifier = Modifier.weight(1.1f)
                                            )
                                            Text(
                                                ourStatusTime,
                                                fontSize = 12.sp,
                                                fontFamily = FontFamily.Monospace,
                                                fontWeight = FontWeight.Bold,
                                                color = if (isOurCanceled) TrainRTheme.RedCancel else TrainRTheme.TextDark,
                                                modifier = Modifier.weight(1.2f)
                                            )
                                            Surface(
                                                color = reconcileColor.copy(alpha = 0.14f),
                                                shape = RoundedCornerShape(6.dp),
                                                border = BorderStroke(1.dp, reconcileColor),
                                                modifier = Modifier.weight(1.3f)
                                            ) {
                                                Text(
                                                    reconcileLabel,
                                                    color = reconcileColor,
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
                        } else {
                            Surface(
                                color = Color(0xFFF8FAFC),
                                shape = RoundedCornerShape(8.dp),
                                border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                                modifier = Modifier.fillMaxSize()
                            ) {
                                Column(modifier = Modifier.fillMaxSize().padding(12.dp)) {
                                    Text(
                                        "Содержимое загруженного PDF файла Яндекс.Расписаний:",
                                        fontWeight = FontWeight.Bold,
                                        fontSize = 13.sp,
                                        color = TrainRTheme.NavyDark
                                    )
                                    Spacer(modifier = Modifier.height(6.dp))
                                    Text(
                                        yandexRawText.ifBlank { "Загруженный PDF-файл не содержит текстового слоя (возможно, это скан-изображение)." },
                                        fontSize = 12.sp,
                                        fontFamily = FontFamily.Monospace,
                                        color = TrainRTheme.TextDark,
                                        modifier = Modifier.fillMaxSize().verticalScroll(rememberScrollState())
                                    )
                                }
                            }
                        }
                    }

                    YandexViewMode.COMPARE_CHANGES -> {
                        if (rows.isEmpty()) {
                            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                Text(
                                    "По станции «$activeStation» нет загруженных изменений из выгрузки или телеграммы для сверки. Переключитесь на «ДАННЫЕ ИЗ PDF ЯНДЕКСА», чтобы видеть все рейсы из PDF.",
                                    fontSize = 13.sp,
                                    color = TrainRTheme.TextSub,
                                    textAlign = TextAlign.Center
                                )
                            }
                        } else {
                            LazyColumn(modifier = Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(8.dp)) {
                                items(rows, key = { it.id }) { row -> DetailedReportTrainCard(row) { onRowClick(row) } }
                            }
                        }
                    }

                    YandexViewMode.RAW_TEXT -> {
                        Surface(
                            color = Color(0xFFF8FAFC),
                            shape = RoundedCornerShape(8.dp),
                            border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                            modifier = Modifier.fillMaxSize()
                        ) {
                            Text(
                                yandexRawText.ifBlank { "Текст в PDF не найден." },
                                fontSize = 12.sp,
                                fontFamily = FontFamily.Monospace,
                                color = TrainRTheme.TextDark,
                                modifier = Modifier.fillMaxSize().padding(14.dp).verticalScroll(rememberScrollState())
                            )
                        }
                    }
                }
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
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column(modifier = Modifier.weight(1f)) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text(
                                "ПОЛНЫЙ ОТЧЕТ ПО ПОЕЗДУ (ДОСЬЕ) — №${trainRow.trainNumber}",
                                fontSize = 17.sp,
                                fontWeight = FontWeight.ExtraBold,
                                color = TrainRTheme.NavyDark
                            )
                            Spacer(modifier = Modifier.width(12.dp))
                            Surface(color = TrainRTheme.RoyalBlue, shape = RoundedCornerShape(6.dp)) {
                                Text(
                                    "СТАНЦИЯ ОТЧЕТА: ${trainRow.station.uppercase()}",
                                    color = Color.White,
                                    fontSize = 12.sp,
                                    fontWeight = FontWeight.ExtraBold,
                                    modifier = Modifier.padding(horizontal = 10.dp, vertical = 4.dp)
                                )
                            }
                            if (trainRow.isFullCancel) {
                                Spacer(modifier = Modifier.width(10.dp))
                                Surface(color = TrainRTheme.RedCancel, shape = RoundedCornerShape(6.dp)) {
                                    Text(
                                        "ПОЕЗД ОТМЕНЕН",
                                        color = Color.White,
                                        fontSize = 12.sp,
                                        fontWeight = FontWeight.ExtraBold,
                                        modifier = Modifier.padding(horizontal = 10.dp, vertical = 4.dp)
                                    )
                                }
                            }
                        }
                        Spacer(modifier = Modifier.height(4.dp))
                        val tgHeader = if (trainRow.telegramExcerpts.isNotEmpty()) {
                            "Найдено в телеграммах: ${trainRow.telegramExcerpts.size} шт. (${trainRow.telegramBasis})"
                        } else {
                            "Телеграмма: не загружена"
                        }
                        Text(
                            "Станция: ${trainRow.station} • Маршрут: ${trainRow.route} • Дата: ${trainRow.dateStr} • $tgHeader",
                            fontSize = 12.sp,
                            color = TrainRTheme.TextSub,
                            fontWeight = FontWeight.SemiBold
                        )
                    }
                    Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RedCancel)) {
                        Text("Закрыть", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                    }
                }
                Spacer(modifier = Modifier.height(8.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                Spacer(modifier = Modifier.height(8.dp))
                Text(
                    "Нитка следования по станциям (отчет открыт для ст. ${trainRow.station}):",
                    fontWeight = FontWeight.Bold,
                    fontSize = 13.sp,
                    color = TrainRTheme.NavyDark
                )
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
                            yandexTime = trainRow.yandexTime,
                            isFullTrainCanceled = trainRow.isFullCancel,
                            isSkippedByExport = trainRow.isCanceled
                        )
                    )
                }

                Column(modifier = Modifier.fillMaxWidth().weight(0.52f).border(1.dp, TrainRTheme.BorderCol, RoundedCornerShape(8.dp))) {
                    Row(modifier = Modifier.fillMaxWidth().background(TrainRTheme.NavyDark).padding(horizontal = 10.dp, vertical = 6.dp)) {
                        Text("СТАНЦИЯ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.6f))
                        Text("АФИША", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(0.9f))
                        Text("ТЕЛЕГРАММА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.2f))
                        Text("ВЫГРУЗКА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.1f))
                        Text("РАЗНИЦА", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(0.9f))
                        Text("СТАТУС ОСТАНОВКИ", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.6f))
                    }
                    LazyColumn(modifier = Modifier.fillMaxWidth()) {
                        items(stopsToShow) { stop ->
                            val isCurrentReportStation = stop.stationName.equals(trainRow.station, ignoreCase = true)
                            val isFullCancelAnywhere = trainRow.isFullCancel || stop.isFullTrainCanceled
                            val isExpCanceled = isFullCancelAnywhere || stop.isSkippedByExport || stop.exportTime.contains("ОТМЕН", true) || stop.exportTime.contains("БЕЗ ОСТ", true)
                            val isTgNoStop = trainRow.hasTelegram && (isFullCancelAnywhere || stop.isSkippedByTelegram || stop.isTechStop || stop.telegramTime.contains("ОТМЕН", true) || stop.telegramTime.contains("БЕЗ ОСТ", true) || stop.telegramTime.contains("ТЕХ", true))

                            val displayExportTime = when {
                                stop.exportTime != "—" -> stop.exportTime
                                isFullCancelAnywhere && trainRow.hasExport -> if (stop.afishaTime != "—") "${stop.afishaTime} ОТМЕНА" else "ОТМЕНА"
                                else -> stop.exportTime
                            }

                            val activeComparisonTime = displayExportTime.takeIf { it != "—" } ?: stop.telegramTime
                            val diffStr = if (isExpCanceled || isTgNoStop) "ОТМЕНА" else calculateMinutesDiff(stop.afishaTime, activeComparisonTime)

                            val stopStatusText = when {
                                isFullCancelAnywhere -> "ОТМЕНЕН"
                                stop.isTechStop -> "ТЕХ. СТОЯНКА (без пасс.)"
                                isExpCanceled || isTgNoStop -> "БЕЗ ОСТАНОВКИ / ОТМЕНА"
                                stop.isAddedByExport || stop.isAddedByTelegram -> "НАЗНАЧЕНА ОСТАНОВКА"
                                stop.isInitialStation -> "Начальная станция (отпр.)"
                                stop.isFinalStation -> "Конечная станция (приб.)"
                                else -> "Посадка / высадка пасс."
                            }

                            val isNegativeStatus = isFullCancelAnywhere || stop.isTechStop || isExpCanceled || isTgNoStop

                            Row(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .background(
                                        when {
                                            isCurrentReportStation -> TrainRTheme.ActionBlue.copy(alpha = 0.12f)
                                            isNegativeStatus -> TrainRTheme.RedCancel.copy(alpha = 0.04f)
                                            else -> Color.Transparent
                                        }
                                    )
                                    .padding(horizontal = 10.dp, vertical = 5.dp),
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Text(
                                    if (isCurrentReportStation) "▶ ${stop.stationName} (ст. отчета)" else stop.stationName,
                                    fontWeight = if (isCurrentReportStation) FontWeight.ExtraBold else FontWeight.Bold,
                                    color = if (isCurrentReportStation) TrainRTheme.RoyalBlue else TrainRTheme.TextDark,
                                    fontSize = 11.sp,
                                    modifier = Modifier.weight(1.6f)
                                )
                                Text(stop.afishaTime, fontFamily = FontFamily.Monospace, fontSize = 11.sp, modifier = Modifier.weight(0.9f))
                                Text(
                                    if (trainRow.hasTelegram) stop.telegramTime else "—",
                                    fontFamily = FontFamily.Monospace,
                                    fontWeight = FontWeight.Bold,
                                    color = if (isTgNoStop) TrainRTheme.RedCancel else TrainRTheme.TextDark,
                                    fontSize = 11.sp,
                                    modifier = Modifier.weight(1.2f)
                                )
                                Text(
                                    displayExportTime,
                                    fontFamily = FontFamily.Monospace,
                                    fontWeight = FontWeight.Bold,
                                    color = if (isExpCanceled) TrainRTheme.RedCancel else TrainRTheme.TextSub,
                                    fontSize = 11.sp,
                                    modifier = Modifier.weight(1.1f)
                                )
                                Text(
                                    diffStr,
                                    fontFamily = FontFamily.Monospace,
                                    fontWeight = if (diffStr == "ОТМЕНА") FontWeight.Bold else FontWeight.Normal,
                                    color = if (diffStr == "ОТМЕНА") TrainRTheme.RedCancel else TrainRTheme.TextDark,
                                    fontSize = 11.sp,
                                    modifier = Modifier.weight(0.9f)
                                )
                                Text(
                                    stopStatusText,
                                    fontSize = 10.sp,
                                    fontWeight = if (isNegativeStatus || stop.isInitialStation || stop.isFinalStation || stop.isAddedByExport || stop.isAddedByTelegram) FontWeight.ExtraBold else FontWeight.Normal,
                                    color = when {
                                        isNegativeStatus -> TrainRTheme.RedCancel
                                        stop.isInitialStation || stop.isFinalStation || stop.isAddedByExport || stop.isAddedByTelegram -> TrainRTheme.GreenSuccess
                                        else -> TrainRTheme.TextDark
                                    },
                                    modifier = Modifier.weight(1.6f)
                                )
                            }
                            HorizontalDivider(color = TrainRTheme.BorderCol.copy(alpha = 0.5f))
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                Row(modifier = Modifier.weight(0.48f).fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(14.dp)) {
                    Column(modifier = Modifier.weight(1f).fillMaxHeight()) {
                        val tgTitle = if (trainRow.telegramExcerpts.isNotEmpty()) {
                            "Все телеграммы, где фигурирует поезд №${trainRow.trainNumber} (${trainRow.telegramExcerpts.size} шт.):"
                        } else {
                            "Выдержка из Телеграммы:"
                        }
                        Text(tgTitle, fontWeight = FontWeight.Bold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
                        Spacer(modifier = Modifier.height(4.dp))
                        Surface(color = Color(0xFFF1F5F9), shape = RoundedCornerShape(8.dp), modifier = Modifier.fillMaxSize()) {
                            if (trainRow.telegramExcerpts.isNotEmpty()) {
                                Column(
                                    modifier = Modifier.fillMaxSize().padding(12.dp).verticalScroll(rememberScrollState()),
                                    verticalArrangement = Arrangement.spacedBy(10.dp)
                                ) {
                                    trainRow.telegramExcerpts.forEachIndexed { idx, item ->
                                        Surface(
                                            color = Color.White,
                                            shape = RoundedCornerShape(6.dp),
                                            border = BorderStroke(1.dp, TrainRTheme.GoldAccent)
                                        ) {
                                            Column(modifier = Modifier.fillMaxWidth().padding(10.dp)) {
                                                Text(
                                                    "Телеграмма #${idx + 1}: ${item.telegramNumber} • от ${item.telegramDate} • ${item.telegramPage}",
                                                    fontWeight = FontWeight.ExtraBold,
                                                    fontSize = 11.sp,
                                                    color = TrainRTheme.RoyalBlue
                                                )
                                                Text(
                                                    "Файл: ${item.fileName}",
                                                    fontSize = 10.sp,
                                                    color = TrainRTheme.TextSub
                                                )
                                                Spacer(modifier = Modifier.height(4.dp))
                                                Text(
                                                    item.snippet,
                                                    fontSize = 11.sp,
                                                    fontFamily = FontFamily.Monospace,
                                                    color = TrainRTheme.TextDark
                                                )
                                            }
                                        }
                                    }
                                }
                            } else {
                                Text(
                                    "Телеграмма для данного поезда не загружена.",
                                    fontSize = 11.sp,
                                    fontFamily = FontFamily.Monospace,
                                    modifier = Modifier.padding(12.dp)
                                )
                            }
                        }
                    }
                    Column(modifier = Modifier.weight(1f).fillMaxHeight()) {
                        Text("Выдержка из Выгрузки по ст. ${trainRow.station} (${trainRow.exportBasis}):", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
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

                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.End) {
                    Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark)) {
                        Text("Закрыть досье", color = Color.White, fontWeight = FontWeight.Bold)
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
                        Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RedCancel)) {
                            Text("Закрыть", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 12.sp)
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))
                HorizontalDivider(color = TrainRTheme.BorderCol)
                Spacer(modifier = Modifier.height(10.dp))

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
                    Button(onClick = onDismiss, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue)) {
                        Text("Применить выбранные (${selectedDates.size})", color = Color.White, fontWeight = FontWeight.Bold)
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
