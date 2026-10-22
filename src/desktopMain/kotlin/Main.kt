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
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
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
import androidx.compose.ui.window.Window
import androidx.compose.ui.window.WindowState
import androidx.compose.ui.window.application
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.Job
import kotlinx.coroutines.delay
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.pdfbox.Loader
import org.apache.pdfbox.text.PDFTextStripper
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.awt.FileDialog
import java.awt.Frame
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.time.YearMonth
import java.util.Calendar

// ==================== КОНСТАНТЫ И МОДЕЛИ ====================

object AppConstants {
    const val DEFAULT_MONTH_NUM = 10
    const val DEFAULT_MONTH_NAME = "октября"
    const val UNGROUPED_DAY_KEY = 999
    const val DEFAULT_STATION_ORDER = 27
    const val TRACK_HINT_LOOKAHEAD = 45
    const val REPARSE_DEBOUNCE_MS = 150L
}

enum class Direction { TO_MOSCOW, FROM_MOSCOW }
enum class TrainFilter { ALL, MCD, EXPRESS }

data class AfishaTrainInfo(
    val trainNum: String, val daysFlag: String, val startStation: String, val endStation: String,
    val routeString: String, val destinationLabel: String, val isMcd: Boolean, val direction: Direction,
    val stationTimes: Map<String, String>, val stationTrackHints: Map<String, String>
)

data class TrackSegmentRule(val fromStationKey: String, val toStationKey: String, val trackNumber: String)
data class CalendarDaySet(val monthNum: Int, val monthName: String, val days: List<Int>, val fallbackLabel: String)
data class SingleDateFilterOption(val day: Int, val monthNum: Int, val monthName: String, val displayLabel: String)

data class TelegramTrainEvent(
    val trainNum: String, val rawDates: String, val calendarDays: CalendarDaySet,
    val telegramHeader: String, val routeFromTelegram: String,
    val isCancelledFull: Boolean, val partialCancelSection: String, val isAppointedFull: Boolean,
    val stationTimes: LinkedHashMap<String, Pair<String, Boolean>>,
    val skippedStops: List<String>, val appointedStops: List<String>,
    val trackRules: List<TrackSegmentRule>, val explicitNoStopsText: String, val rawBlockText: String,
    val isClarificationNotice: Boolean = false, val clarificationRouteNotice: String = ""
)

data class StationComparisonRow(
    val dateGroup: String, val sortDayKey: Int, val activeDays: List<Int>, val monthName: String,
    val trainNum: String, val daysFlag: String, val routeString: String, val destinationLabel: String,
    val isMcd: Boolean, val direction: Direction, val stationName: String, val telegramHeader: String,
    val afishaTime: String, val telegramTime: String, val otmenyFileStatus: String, val finalChangeText: String,
    val isCancelledAtStation: Boolean, val isPassWithoutStop: Boolean, val isAppointedAtStation: Boolean, val isTimeShiftedAtStation: Boolean,
    val isPartialCancel: Boolean, val isAppointedStop: Boolean,
    val skippedStopsSummary: String, val appointedStopsSummary: String, val trackSummaryForStation: String, val hasDiscrepancyWithOtmeny: Boolean,
    val fullAfishaStations: Map<String, String>, val fullTelegramStations: Map<String, Pair<String, Boolean>>, val rawTelegramText: String,
    val validationError: String = ""
)

data class YandexParsedTrain(
    val trainNum: String,
    val departureTime: String,
    val route: String,
    val stopsRule: String,
    val rawText: String
)

enum class YandexDiscrepancySeverity { CRITICAL_CANCEL_MISSED, TIME_MISMATCH, STOP_MISMATCH, INFO }

data class YandexDiscrepancyRow(
    val trainNum: String,
    val route: String,
    val direction: Direction,
    val yandexTime: String,
    val actualExpectedTime: String,
    val yandexStops: String,
    val telegramStopsChange: String,
    val basisDocument: String,
    val discrepancyDescription: String,
    val severity: YandexDiscrepancySeverity
)

data class GlobalError(val id: String, val trainNum: String, val header: String, val reason: String, var isIgnored: Boolean = false)

// ==================== ТЕМА И СТАНЦИИ ====================

object TrainRTheme {
    val BgMain = Color(0xFFF1F5F9)
    val CardBg = Color(0xFFFFFFFF)
    val NavyDark = Color(0xFF0F172A)
    val RoyalBlue = Color(0xFF1E3A8A)
    val McdCoral = Color(0xFFBE123C)
    val GoldAccent = Color(0xFFD97706)
    val GreenSuccess = Color(0xFF059669)
    val RedCancel = Color(0xFFDC2626)
    val AmberWarn = Color(0xFFB45309)
    val BorderCol = Color(0xFFCBD5E1)
    val TextDark = Color(0xFF0F172A)
    val TextSub = Color(0xFF334155)
    val PurpleMcd = Color(0xFF7C3AED)
    val YandexRed = Color(0xFFE11D48)
}

val POLYGON_STATIONS = listOf(
    "Сходня", "Химки", "Зеленоград-Крюково", "Подрезково", "Новоподрезково",
    "Молжаниново", "Левобережная", "Алабушево", "Фирсановская", "Ховрино",
    "Грачёвская", "Моссельмаш", "Лихоборы", "Останкино", "Поварово-1",
    "Подсолнечная", "Клин", "Решетниково", "Завидово", "Редкино", "Тверь", "Конаково ГРЭС"
)

val STATION_ORDER_INDEX = mapOf(
    "тверь" to 1, "чуприяновка" to 2, "кузьминка" to 3, "межево" to 4, "редкино" to 5,
    "московскоеморе" to 6, "завидово" to 7, "черничная" to 8, "конаковогрэс" to 6,
    "донховка" to 7, "конаковскиймох" to 8, "путепроводная" to 9, "решетниково" to 10,
    "ямуга" to 11, "клин" to 12, "стреглово" to 13, "фроловское" to 14, "покровка" to 15,
    "головково" to 16, "сенеж" to 17, "подсолнечная" to 18, "березкидачные" to 19,
    "поварово1" to 20, "поваровка" to 21, "радищево" to 22, "алабушево" to 23,
    "крюково" to 24, "малино" to 25, "фирсановская" to 26, "сходня" to 27,
    "подрезково" to 28, "новоподрезково" to 29, "молжаниново" to 30, "химки" to 31,
    "левобережная" to 32, "ховрино" to 33, "грачевская" to 34, "моссельмаш" to 35,
    "лихоборы" to 36, "петровскоразумовская" to 37, "останкино" to 38, "москватоварная" to 39,
    "рижская" to 40, "пострижский" to 41, "москва" to 42, "ипподром" to 43
)

val HARD_CLOSED_STATIONS = setOf("петровскоразумовская", "петрразумовская", "рижская", "пострижский", "малино")

val MONTH_NAMES_BY_NUM = mapOf(
    1 to "января", 2 to "февраля", 3 to "марта", 4 to "апреля", 5 to "мая", 6 to "июня",
    7 to "июля", 8 to "августа", 9 to "сентября", 10 to "октября", 11 to "ноября", 12 to "декабря"
)

val MONTH_NAMES_NOMINATIVE = mapOf(
    1 to "Январь", 2 to "Февраль", 3 to "Март", 4 to "Апрель", 5 to "Май", 6 to "Июнь",
    7 to "Июль", 8 to "Август", 9 to "Сентябрь", 10 to "Октябрь", 11 to "Ноябрь", 12 to "Декабрь"
)

// ==================== БАЗОВЫЕ УТИЛИТЫ И ПАРСЕРЫ ====================

fun getCachedAfishaFile(): File {
    val dir = File(System.getProperty("user.home"), ".train_r_system")
    if (!dir.exists()) dir.mkdirs()
    return File(dir, "cached_afisha.xlsx")
}

private fun logWarn(msg: String, e: Throwable? = null) {
    System.err.println("[TrainR/WARN] $msg")
    e?.printStackTrace()
}

fun readDocumentText(file: File): String {
    return try {
        when {
            file.name.endsWith(".docx", true) -> {
                FileInputStream(file).use { fis ->
                    XWPFDocument(fis).use { doc ->
                        doc.paragraphs.joinToString("\n") { it.text }
                    }
                }
            }
            file.name.endsWith(".pdf", true) -> {
                Loader.loadPDF(file).use { doc -> PDFTextStripper().getText(doc) }
            }
            else -> file.readText(Charsets.UTF_8)
        }
    } catch (e: Exception) {
        logWarn("Ошибка чтения файла ${file.name}", e)
        ""
    }
}

fun preprocessTelegramLine(rawLine: String): String {
    var s = rawLine
    s = s.replace(Regex("(?i)(нр\\.?|поезд|№)(?=\\d)"), "$1 ")
    s = s.replace(Regex("(?<![:.])(\\b\\d{4}(?:/\\d{4})?)(?=[А-Яа-яЁё])"), "$1 ")
    s = s.replace(Regex("([А-Яа-яЁё])(?=\\d{1,2}[:.]\\d{2})"), "$1 ")
    s = s.replace(Regex("([А-Яа-яЁё0-9])\\("), "$1 (")
    return s.replace(Regex("\\s+"), " ").trim()
}

fun isStartOfNewTrainLine(line: String, inCancelBlock: Boolean, inAppointBlock: Boolean): Pair<Boolean, String> {
    val explicitMatch = Regex("(?i)^(?:нр\\.?|поезд\\s*№?|№)\\s*:?\\s*([67]\\d{3}(?:/[67]\\d{3})?)\\b").find(line)
    if (explicitMatch != null) return Pair(true, explicitMatch.groupValues[1])
    val bareMatch = Regex("^([67]\\d{3}(?:/[67]\\d{3})?)(?![:.\\d])\\s+([А-Яа-яЁё].*)").find(line)
    if (bareMatch != null) {
        val rest = bareMatch.groupValues[2]
        if (inCancelBlock || inAppointBlock || rest.contains("отмен", true) || rest.contains("назнач", true) || Regex("[-–—]").containsMatchIn(rest.take(35))) {
            return Pair(true, bareMatch.groupValues[1])
        }
    }
    return Pair(false, "")
}

fun splitDirectionalTrainNumbers(trainNum: String): List<String> = trainNum.split("/").map { it.trim() }.filter { it.isNotEmpty() }.distinct()

fun trainNumLookupKeys(trainNum: String): Set<String> {
    val parts = splitDirectionalTrainNumbers(trainNum)
    val keys = parts.toMutableSet()
    if (parts.size >= 2) { keys.add("${parts[0]}/${parts[1]}"); keys.add("${parts[1]}/${parts[0]}") }
    return keys
}

fun getDirectionBySingleNumber(singleNum: String): Direction {
    val digits = singleNum.filter { it.isDigit() }
    val lastDigit = digits.lastOrNull()?.digitToIntOrNull() ?: 1
    return if (lastDigit % 2 == 0) Direction.FROM_MOSCOW else Direction.TO_MOSCOW
}

fun isTrainMcd(num: String): Boolean {
    val base = num.split("/").first().filter { it.isDigit() }
    return "/" in num || base.startsWith("72") || base.startsWith("73") || base.startsWith("74") || base.startsWith("67")
}

fun isTrainExpress(num: String): Boolean {
    val base = num.split("/").first().filter { it.isDigit() }
    return base.startsWith("7") && base.length == 4 && !isTrainMcd(num)
}

fun normalizeStationKey(name: String): String {
    val lower = name.lowercase().replace("ё", "е").replace(Regex("\\(закрыта\\)"), "").replace(Regex("\\(47\\s*км\\)"), "").replace(Regex("\\(ст\\.?\\s*ховрино\\)"), "").trim()
    if (Regex("\\b(оп|о\\.п\\.|пл|пл\\.)\\s*ховрин").containsMatchIn(lower)) return "грачевская"
    val clean = lower.replace(Regex("(?i)\\b(оп|о\\.п\\.|рп|р\\.п\\.|ст|ст\\.|пл|пл\\.|ост\\.?\\s*пункт)\\b"), "").replace(Regex("[^а-я0-9]"), "")
    if (clean.isEmpty()) return ""
    return when {
        "крюков" in clean || "зеленогр" in clean || "зелкрюк" in clean -> "крюково"
        "петровско" in clean || "петрразум" in clean || "празумов" in clean || "петрраз" in clean -> "петровскоразумовская"
        "грачевск" in clean -> "грачевская"
        "ховрин" in clean -> "ховрино"
        "новоподрез" in clean || "нподрез" in clean -> "новоподрезково"
        "подрезков" in clean -> "подрезково"
        "молжанин" in clean -> "молжаниново"
        "левобереж" in clean || "лбережн" in clean -> "левобережная"
        "фирсанов" in clean -> "фирсановская"
        "малин" in clean -> "малино"
        "сходн" in clean -> "сходня"
        "химк" in clean -> "химки"
        "моссельмаш" in clean || "моссельм" in clean -> "моссельмаш"
        "лихобор" in clean -> "лихоборы"
        "останкин" in clean -> "останкино"
        "постриж" in clean -> "пострижский"
        "рижск" in clean -> "рижская"
        ("моск" in clean || "мск" in clean) && "тов" in clean -> "москватоварная"
        "моск" in clean || "мск" in clean || "ленингр" in clean || "лвок" in clean -> "москва"
        "ипподр" in clean || "47км" in clean || "раменск" in clean -> "ипподром"
        "алабуш" in clean -> "алабушево"
        "радищ" in clean -> "радищево"
        "поваровк" in clean -> "поваровка"
        "поваров" in clean -> "поварово1"
        "березк" in clean -> "березкидачные"
        "подсолн" in clean -> "подсолнечная"
        "сенеж" in clean -> "сенеж"
        "головк" in clean -> "головково"
        "покровк" in clean -> "покровка"
        "фролов" in clean -> "фроловское"
        "стрегл" in clean || "стригл" in clean -> "стреглово"
        "клин" in clean -> "клин"
        "ямуг" in clean -> "ямуга"
        "решетник" in clean -> "решетниково"
        "завидов" in clean -> "завидово"
        "черничн" in clean -> "черничная"
        "москморе" in clean || "московскоеморе" in clean -> "московскоеморе"
        "редкин" in clean -> "редкино"
        "межев" in clean -> "межево"
        "кузьмин" in clean -> "кузьминка"
        "чуприян" in clean -> "чуприяновка"
        "твер" in clean -> "тверь"
        "конаков" in clean && "мох" in clean -> "конаковскиймох"
        "донхов" in clean -> "донховка"
        "путепров" in clean -> "путепроводная"
        "конаков" in clean -> "конаковогрэс"
        else -> clean
    }
}

fun displayStationName(normKey: String): String {
    return when (normKey) {
        "крюково" -> "Зеленоград-Крюково"; "грачевская" -> "Грачёвская (ст. Ховрино)"; "ховрино" -> "Ховрино"; "новоподрезково" -> "Новоподрезково"
        "подрезково" -> "Подрезково"; "молжаниново" -> "Молжаниново"; "левобережная" -> "Левобережная"; "фирсановская" -> "Фирсановская"
        "малино" -> "Малино"; "сходня" -> "Сходня"; "химки" -> "Химки"; "моссельмаш" -> "Моссельмаш"; "лихоборы" -> "Лихоборы"
        "останкино" -> "Останкино"; "петровскоразумовская" -> "Петровско-Разумовская"; "москватоварная" -> "Москва-Товарная"
        "москва" -> "Москва"; "ипподром" -> "Ипподром"; "алабушево" -> "Алабушево"; "радищево" -> "Радищево"; "поваровка" -> "Поваровка"
        "поварово1" -> "Поварово-1"; "березкидачные" -> "Березки-Дачные"; "подсолнечная" -> "Подсолнечная"; "сенеж" -> "Сенеж"
        "головково" -> "Головково"; "покровка" -> "Покровка"; "фроловское" -> "Фроловское"; "стреглово" -> "Стреглово"
        "клин" -> "Клин"; "ямуга" -> "Ямуга"; "решетниково" -> "Решетниково"; "черничная" -> "Черничная"; "завидово" -> "Завидово"
        "московскоеморе" -> "Московское Море"; "редкино" -> "Редкино"; "межево" -> "Межево"; "кузьминка" -> "Кузьминка"
        "чуприяновка" -> "Чуприяновка"; "тверь" -> "Тверь"; "путепроводная" -> "Путепроводная"; "конаковскиймох" -> "Конаковский Мох"
        "донховка" -> "Донховка"; "конаковогрэс" -> "Конаково ГРЭС"
        else -> normKey.replaceFirstChar { it.uppercase() }
    }
}

fun cleanSingleTimeToken(raw: String): String {
    if (raw.isBlank()) return ""
    var s = raw.split("@")[0].trim()
    if ("/" in s) s = s.split("/").last().trim()
    s = s.replace(Regex("\\([+\\-]\\d+\\)"), "").trim().split(",")[0].trim().replace(".", ":")
    val nums = Regex("\\d+").findAll(s).map { it.value }.toList()
    if (nums.size >= 2) {
        val h = nums[0].toIntOrNull() ?: return ""
        val m = nums[1].toIntOrNull() ?: return ""
        if (h in 0..23 && m in 0..59) return "%02d:%02d".format(h, m)
    }
    return ""
}

fun calculateMinutesDiffLabel(afishaTime: String, telegramTime: String): String {
    val p1 = afishaTime.split(":"); val p2 = telegramTime.split(":")
    if (p1.size < 2 || p2.size < 2) return ""
    val m1 = (p1[0].toIntOrNull() ?: return "") * 60 + (p1[1].toIntOrNull() ?: return "")
    val m2 = (p2[0].toIntOrNull() ?: return "") * 60 + (p2[1].toIntOrNull() ?: return "")
    var diff = m2 - m1
    if (diff < -720) diff += 1440; if (diff > 720) diff -= 1440
    return when { diff > 0 -> "(+$diff мин)"; diff < 0 -> "($diff мин)"; else -> "(без сдвига)" }
}

fun detectMonthNumFromString(s: String, defaultMonth: Int = AppConstants.DEFAULT_MONTH_NUM): Int {
    val low = s.lowercase()
    return when {
        "янв" in low -> 1; "фев" in low -> 2; "мар" in low -> 3; "апр" in low -> 4; "мая" in low || "май" in low -> 5
        "июн" in low -> 6; "июл" in low -> 7; "авг" in low -> 8; "сен" in low -> 9; "окт" in low -> 10; "ноя" in low -> 11; "дек" in low -> 12
        else -> { val slashM = Regex("\\d{1,2}[/.\\-](\\d{1,2})").find(low)?.groupValues?.get(1)?.toIntOrNull(); if (slashM != null && slashM in 1..12) slashM else defaultMonth }
    }
}

fun parseCalendarDays(raw: String, fallbackRaw: String): CalendarDaySet {
    val source = if (raw.isNotBlank()) raw else fallbackRaw
    val clean = source.replace("(", "").replace(")", "").trim()
    val mNum = detectMonthNumFromString(clean, detectMonthNumFromString(fallbackRaw, AppConstants.DEFAULT_MONTH_NUM))
    val mName = MONTH_NAMES_BY_NUM[mNum] ?: AppConstants.DEFAULT_MONTH_NAME
    val daysSet = sortedSetOf<Int>()
    val rangeRegex = Regex("(?i)с\\s*(\\d{1,2})(?:[/.\\-]\\d{1,2})?(?:\\s*[а-я]+)?\\s*по\\s*(\\d{1,2})(?:[/.\\-]\\d{1,2})?")
    for (rm in rangeRegex.findAll(clean)) { val d1 = rm.groupValues[1].toIntOrNull() ?: continue; val d2 = rm.groupValues[2].toIntOrNull() ?: continue; if (d1 in 1..31 && d2 in 1..31 && d1 <= d2) for (d in d1..d2) daysSet.add(d) }
    val listSlashRegex = Regex("((?:\\d{1,2}\\s*,\\s*)*\\d{1,2})[/.\\-](\\d{1,2})")
    for (lm in listSlashRegex.findAll(clean)) {
        val prefix = clean.substring(0, lm.range.first)
        if (prefix.trimEnd().endsWith("с", true) || prefix.trimEnd().endsWith("по", true) || prefix.trimEnd().endsWith("кроме", true) || prefix.trimEnd().endsWith("кр.", true)) continue
        lm.groupValues[1].split(",").forEach { p -> val d = p.trim().toIntOrNull(); if (d != null && d in 1..31) daysSet.add(d) }
    }
    val onlyMatch = Regex("(?i)только\\s*(\\d{1,2})").find(clean); if (onlyMatch != null) { val d = onlyMatch.groupValues[1].toIntOrNull(); if (d != null && d in 1..31) daysSet.add(d) }
    if (daysSet.isEmpty()) { val wordListMatch = Regex("((?:\\d{1,2}\\s*,\\s*)*\\d{1,2})\\s+[а-я]{3,}").find(clean); if (wordListMatch != null) wordListMatch.groupValues[1].split(",").forEach { p -> val d = p.trim().toIntOrNull(); if (d != null && d in 1..31) daysSet.add(d) } }
    val exceptRegex = Regex("(?i)(?:кроме|кр\\.?)\\s*(\\d{1,2})"); for (em in exceptRegex.findAll(clean)) { val exDay = em.groupValues[1].toIntOrNull(); if (exDay != null) daysSet.remove(exDay) }
    return CalendarDaySet(mNum, mName, daysSet.toList(), clean.ifEmpty { "По датам телеграммы" })
}

fun formatDaysBlockHeader(days: List<Int>, monthName: String): String {
    if (days.isEmpty()) return "По датам телеграммы"
    val sorted = days.sorted(); val isContiguous = (sorted.last() - sorted.first() + 1 == sorted.size)
    return when { sorted.size == 1 -> "${sorted.first()} $monthName"; isContiguous && sorted.size >= 3 -> "с ${sorted.first()} по ${sorted.last()} $monthName"; else -> "${sorted.joinToString(", ")} $monthName" }
}

fun extractAvailableDateOptions(events: List<TelegramTrainEvent>): List<SingleDateFilterOption> {
    val map = mutableMapOf<Pair<Int, Int>, SingleDateFilterOption>()
    for (ev in events) for (d in ev.calendarDays.days) {
        val key = Pair(ev.calendarDays.monthNum, d)
        if (!map.containsKey(key)) map[key] = SingleDateFilterOption(d, ev.calendarDays.monthNum, ev.calendarDays.monthName, "$d ${ev.calendarDays.monthName}")
    }
    return map.values.sortedWith(compareBy({ it.monthNum }, { it.day }))
}

fun formatTrackBadgeForStation(stationNorm: String, trackNum: String, rawHint: String = ""): String {
    if (stationNorm != "сходня" && stationNorm != "химки") return ""
    if (rawHint.isNotEmpty()) return "(по ${rawHint.removePrefix("@").trim()})"
    return when (trackNum) {
        "1" -> "(по 1 пл. 1 пути)"; "2" -> "(по 2 пл. 2 пути)"; "3" -> if (stationNorm == "сходня") "(по 3 пл. 3 пути)" else "(по 3 пути)"
        "4" -> if (stationNorm == "химки") "(по 3 пл. 4 пути)" else "(по 4 пути)"; else -> "(по $trackNum пути)"
    }
}

fun resolveTrackForStation(stationNorm: String, trainNum: String, direction: Direction, rules: List<TrackSegmentRule>, afishaHint: String): String {
    if (stationNorm != "сходня" && stationNorm != "химки") return ""
    if (afishaHint.isNotEmpty()) return formatTrackBadgeForStation(stationNorm, "", afishaHint)
    val stOrder = STATION_ORDER_INDEX[stationNorm] ?: return ""
    for (rule in rules) {
        val idx1 = STATION_ORDER_INDEX[rule.fromStationKey] ?: continue; val idx2 = STATION_ORDER_INDEX[rule.toStationKey] ?: continue
        val minIdx = minOf(idx1, idx2); val maxIdx = maxOf(idx1, idx2)
        if (stOrder in minIdx..maxIdx) {
            val isMovingTowardsMoscow = (direction == Direction.TO_MOSCOW)
            val isEndBoundary = (isMovingTowardsMoscow && stOrder == maxIdx && minIdx < maxIdx) || (!isMovingTowardsMoscow && stOrder == minIdx && minIdx < maxIdx)
            if (isEndBoundary && rules.size > 1) continue
            return formatTrackBadgeForStation(stationNorm, rule.trackNumber)
        }
    }
    return ""
}

fun lookupOtmenyStatus(otmenyMap: Map<String, String>, directionalNum: String): String {
    otmenyMap[directionalNum]?.let { return it }
    return otmenyMap.entries.firstOrNull { (k, _) -> directionalNum in trainNumLookupKeys(k) }?.value ?: ""
}

fun lookupAfisha(afishaDb: Map<String, AfishaTrainInfo>, directionalNum: String): AfishaTrainInfo? {
    afishaDb[directionalNum]?.let { return it }
    return afishaDb.entries.firstOrNull { (k, _) -> directionalNum in trainNumLookupKeys(k) }?.value
}

// ==================== ПАРСЕР БАЗОВОЙ АФИШИ (EXCEL) ====================

fun loadAfishaDatabase(afishaFile: File?): Map<String, AfishaTrainInfo> {
    if (afishaFile == null || !afishaFile.exists()) return emptyMap()
    val db = mutableMapOf<String, AfishaTrainInfo>()
    val trackHintsDb = mutableMapOf<String, MutableMap<String, String>>()
    val timeFmt = SimpleDateFormat("HH:mm")
    try {
        FileInputStream(afishaFile).use { fis ->
            val wb = WorkbookFactory.create(fis)
            for (sIdx in 0 until wb.numberOfSheets) {
                val sheet = wb.getSheetAt(sIdx)
                var headerRowIdx = 0
                for (r in 0..10) {
                    val row = sheet.getRow(r) ?: continue; var foundTrains = 0
                    for (c in 1..minOf(50, row.lastCellNum.toInt())) { val v = row.getCell(c)?.toString()?.trim()?.replace(".0", "") ?: ""; if (Regex("\\b\\d{4}\\b").containsMatchIn(v)) foundTrains++ }
                    if (foundTrains >= 3) { headerRowIdx = r; break }
                }
                val stationRows = mutableMapOf<Int, String>()
                for (r in (headerRowIdx + 1)..sheet.lastRowNum) {
                    val cellVal = sheet.getRow(r)?.getCell(0)?.toString()?.trim() ?: ""
                    if (cellVal.isNotEmpty() && !cellVal.contains("Условные", true) && !cellVal.contains("Расписание", true) && !cellVal.equals("изм", true) && !cellVal.equals("отм", true)) stationRows[r] = cellVal
                }
                val hRow = sheet.getRow(headerRowIdx) ?: continue
                val subRow = sheet.getRow(headerRowIdx + 1)
                val subRowIsModifiers = subRow?.getCell(0)?.toString()?.trim().isNullOrEmpty()
                for (c in 1..hRow.lastCellNum) {
                    val v1 = hRow.getCell(c)?.toString()?.trim()?.replace(".0", "") ?: ""
                    val v2 = if (subRowIsModifiers) (subRow?.getCell(c)?.toString()?.trim() ?: "") else ""
                    val combined = "$v1 $v2".trim()
                    val numMatch = Regex("\\b(\\d{4}(?:/\\d{4})?)\\b").find(combined) ?: continue
                    val trainNum = numMatch.groupValues[1]
                    val isNormColumn = combined.contains("норм", true) || (!combined.contains("изм", true) && !combined.contains("отм", true))
                    for ((rIdx, rawStName) in stationRows) {
                        val rawStr = sheet.getRow(rIdx)?.getCell(c)?.toString()?.trim() ?: ""
                        if ("@" in rawStr) {
                            val normSt = normalizeStationKey(rawStName)
                            val hint = rawStr.substringAfter("@").trim()
                            var hints = trackHintsDb[trainNum]
                            if (hints == null) {
                                hints = mutableMapOf()
                                trackHintsDb[trainNum] = hints
                            }
                            hints[normSt] = hint
                        }
                    }
                    if (db.containsKey(trainNum) && !isNormColumn) continue
                    val daysFlag = when { combined.contains("СВ", true) -> "СВ"; combined.contains("Р", true) -> "Р"; else -> "" }
                    val stTimes = LinkedHashMap<String, String>(); var firstStRaw = ""; var lastStRaw = ""
                    for ((rIdx, rawStName) in stationRows) {
                        val cell = sheet.getRow(rIdx)?.getCell(c) ?: continue
                        val rawTimeStr = if (cell.cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) timeFmt.format(cell.dateCellValue) else cell.toString().trim()
                        if (rawTimeStr.contains("Конаково", true)) continue
                        val cleaned = cleanSingleTimeToken(rawTimeStr)
                        if (cleaned.isNotEmpty() && !rawStName.contains("закрыта", true)) {
                            val normKey = normalizeStationKey(rawStName); stTimes[normKey] = cleaned
                            if (firstStRaw.isEmpty()) firstStRaw = rawStName; lastStRaw = rawStName
                        }
                    }
                    if (stTimes.isNotEmpty()) {
                        val startNorm = normalizeStationKey(firstStRaw); val endNorm = normalizeStationKey(lastStRaw)
                        val isMcd = startNorm == "пострижский" || endNorm == "пострижский" || isTrainMcd(trainNum)
                        val cleanStart = when (startNorm) { "пострижский", "ипподром" -> "Ипподром"; "москва" -> "Москва"; else -> displayStationName(startNorm) }
                        val cleanEnd = when (endNorm) { "пострижский", "ипподром" -> "Ипподром"; "москва" -> "Москва"; else -> displayStationName(endNorm) }
                        val dir = getDirectionBySingleNumber(splitDirectionalTrainNumbers(trainNum).first())
                        val destLabel = if (dir == Direction.TO_MOSCOW) {
                            if (isMcd || cleanEnd == "Ипподром") "до Ипподрома (47 км)" else "до Москвы"
                        } else {
                            when (endNorm) {
                                "крюково" -> "до Крюково"; "клин" -> "до Клина"; "тверь" -> "до Твери"; "конаковогрэс" -> "до Конаково ГРЭС"
                                "подсолнечная" -> "до Подсолнечной"; "решетниково" -> "до Решетниково"; "поварово1" -> "до Поварово-1"; "поваровка" -> "до Поваровки"; else -> "до $cleanEnd"
                            }
                        }
                        db[trainNum] = AfishaTrainInfo(trainNum, daysFlag, cleanStart, cleanEnd, "$cleanStart – $cleanEnd", destLabel, isMcd, dir, stTimes, trackHintsDb[trainNum] ?: emptyMap())
                    }
                }
            }
        }
        for ((tNum, hints) in trackHintsDb) { db[tNum]?.let { db[tNum] = it.copy(stationTrackHints = hints) } }
    } catch (e: Exception) { logWarn("Не удалось загрузить Афишу", e) }
    return db
}

// ==================== ПАРСЕР ТЕЛЕГРАММ И РАЗЪЯСНЕНИЙ ====================

fun parseClarificationTelegramEvents(
    rawText: String,
    telHeader: String,
    defaultDates: String
): Pair<List<TelegramTrainEvent>, List<GlobalError>> {
    val events = mutableListOf<TelegramTrainEvent>()
    val errors = mutableListOf<GlobalError>()

    val lines = rawText.lines().map { preprocessTelegramLine(it) }
    for (line in lines) {
        if (!line.startsWith("П.", true) && !line.startsWith("п.") && !line.contains("поезд", true)) continue

        val trainNums = Regex("\\b([67]\\d{3}(?:/[67]\\d{3})?)\\b").findAll(line).map { it.value }.toList()
        if (trainNums.isEmpty()) continue

        val allParens = Regex("\\(([^)]+)\\)").findAll(line).map { it.groupValues[1] }.toList()
        val dateParen = allParens.firstOrNull { Regex("\\d{1,2}[/.\\-]\\d{1,2}|нояб|окт").containsMatchIn(it.lowercase()) }
        val inlineDate = Regex("(?i)(?:только\\s*\\d{1,2}[/.\\-]\\d{1,2}|с\\s*\\d{1,2}[/.\\-]\\d{1,2}\\s*по\\s*\\d{1,2}[/.\\-]\\d{1,2}[^)]*|\\d{1,2}(?:,\\s*\\d{1,2})+/[01]?\\d)").find(line)?.value
        val rawDates = dateParen ?: inlineDate ?: defaultDates
        val calDays = parseCalendarDays(rawDates, defaultDates)

        val isStopInsteadOfTech = line.contains("посадка, высадка пассажиров (вместо", true) ||
                line.contains("посадка и высадка", true) || line.contains("посадка, высадка", true)

        val isTechInsteadOfStop = line.contains("стоянка техническая (вместо", true) ||
                line.contains("вместо посадки", true)

        val isNoPassengers = line.contains("без пассажиров", true)

        for (num in trainNums) {
            val stMap = LinkedHashMap<String, Pair<String, Boolean>>()
            val skipped = mutableListOf<String>()
            val appointed = mutableListOf<String>()
            var routeNotice = ""

            if (isStopInsteadOfTech) {
                val mSt = Regex("по\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+посадка", RegexOption.IGNORE_CASE).find(line)
                val stName = mSt?.groupValues?.get(1)?.trim() ?: ""
                val normSt = normalizeStationKey(stName)
                if (normSt.isNotEmpty()) {
                    appointed.add(displayStationName(normSt))
                    stMap[normSt] = Pair("", true)
                }
            } else if (isTechInsteadOfStop) {
                val mSt = Regex("по\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+стоянка\\s+техническая", RegexOption.IGNORE_CASE).find(line)
                val stName = mSt?.groupValues?.get(1)?.trim() ?: ""
                val normSt = normalizeStationKey(stName)
                if (normSt.isNotEmpty()) {
                    skipped.add(displayStationName(normSt))
                    stMap[normSt] = Pair("", false)
                }
            } else if (isNoPassengers) {
                val mDrop = Regex("по\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+высадка\\s+пассажиров.*?далее\\s+до\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+без\\s+пассажиров", RegexOption.IGNORE_CASE).find(line)
                val mSec = Regex("от\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+до\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+без\\s+пассажиров", RegexOption.IGNORE_CASE).find(line)
                val mDep = Regex("([А-Яа-яЁё0-9\\s.\\-]+?)\\s+отправление\\s+без\\s+пассажиров", RegexOption.IGNORE_CASE).find(line)

                if (mDrop != null) {
                    val lastSt = displayStationName(normalizeStationKey(mDrop.groupValues[1]))
                    routeNotice = "следует только до $lastSt"
                } else if (mSec != null) {
                    val lastSt = displayStationName(normalizeStationKey(mSec.groupValues[1]))
                    routeNotice = "следует только до $lastSt"
                } else if (mDep != null) {
                    routeNotice = "посадки нет"
                }
            }

            events.add(
                TelegramTrainEvent(
                    trainNum = num,
                    rawDates = rawDates,
                    calendarDays = calDays,
                    telegramHeader = telHeader,
                    routeFromTelegram = "",
                    isCancelledFull = false,
                    partialCancelSection = "",
                    isAppointedFull = false,
                    stationTimes = stMap,
                    skippedStops = skipped,
                    appointedStops = appointed,
                    trackRules = emptyList(),
                    explicitNoStopsText = "",
                    rawBlockText = line,
                    isClarificationNotice = true,
                    clarificationRouteNotice = routeNotice
                )
            )
        }
    }

    return Pair(events, errors)
}

fun parseTelegramDocument(file: File): Pair<List<TelegramTrainEvent>, List<GlobalError>> {
    val rawText = readDocumentText(file)
    if (rawText.isBlank()) return Pair(emptyList(), emptyList())

    val headerMatch = Regex("(?i)ТЕЛЕГРАММА\\s*№?\\s*([^\\s]+)\\s+от\\s+([0-9/.\\-]+(?:г\\.?)?)").find(rawText)
    val telHeader = if (headerMatch != null) "Тел. №${headerMatch.groupValues[1]} от ${headerMatch.groupValues[2]}" else "Тел. (${file.nameWithoutExtension.take(18)})"
    val globalDatesMatch = Regex("(?i)(?:с\\s*\\d{1,2}\\s+[а-я]+\\s+по\\s+\\d{1,2}\\s+[а-я]+(?:,\\s*с\\s*\\d{1,2}\\s+[а-я]+\\s+по\\s*\\d{1,2}\\s+[а-я]+)?)").find(rawText)
    val defaultDateGroup = globalDatesMatch?.value?.trim() ?: "По датам телеграммы"

    if (rawText.contains("Внести разъяснение", true) || rawText.contains("разъяснение в телеграмм", true)) {
        return parseClarificationTelegramEvents(rawText, telHeader, defaultDateGroup)
    }

    val events = mutableListOf<TelegramTrainEvent>()
    val errors = mutableListOf<GlobalError>()
    val lines = rawText.lines().map { preprocessTelegramLine(it) }
    var inCancelBlock = false; var inAppointBlock = false; var inBologoeSection = false; var currentSectionDates = ""
    var i = 0
    while (i < lines.size) {
        val line = lines[i]; if (line.isEmpty()) { i++; continue }
        if (line.contains("Бологое", true) && (line.contains("участке", true) || line.contains("Окуловка", true))) inBologoeSection = true
        else if (line.contains("пригородных поездов МЦД", true) || line.contains("пригородных поездов:", true) || line.contains("«Ласточка»", true)) inBologoeSection = false
        if (Regex("(?i)(?:^|\\s|[0-9.]+)(?:отменить|не назначать)").containsMatchIn(line) && (line.contains(":") || line.contains("("))) {
            inCancelBlock = true; inAppointBlock = false; val secDate = Regex("\\(([^)]+)\\)").find(line)?.groupValues?.get(1); if (secDate != null) currentSectionDates = secDate; i++; continue
        } else if (Regex("(?i)(?:^|\\s|[0-9.]+)назначить\\s*:").containsMatchIn(line)) {
            inCancelBlock = false; inAppointBlock = true; currentSectionDates = ""; i++; continue
        } else if (Regex("(?i)(?:^|\\s|[0-9.]+)изменить(?:\\s+расписание)?\\s*:").containsMatchIn(line)) {
            inCancelBlock = false; inAppointBlock = false; val secDate = Regex("\\(([^)]+)\\)").find(line)?.groupValues?.get(1); if (secDate != null) currentSectionDates = secDate; i++; continue
        } else if (line.startsWith("Обороты:", true)) { inCancelBlock = false; inAppointBlock = false; i++; continue }
        
        val (isTrainStart, trainNum) = isStartOfNewTrainLine(line, inCancelBlock, inAppointBlock)
        if (isTrainStart && !inBologoeSection) {
            val blockLines = mutableListOf(line); var j = i + 1
            while (j < lines.size) {
                val nl = lines[j]; if (nl.isEmpty()) { j++; continue }
                val (nextIsTrain, _) = isStartOfNewTrainLine(nl, inCancelBlock, inAppointBlock)
                if (nextIsTrain || Regex("(?i)^(?:нр\\.?|поезд\\s*№?)\\s*\\d{1,3}\\b").containsMatchIn(nl) || Regex("(?i)(?:^|\\s|[0-9.]+)(?:отменить|не назначать|назначить|изменить(?:\\s+расписание)?|обороты)\\s*[:(]").containsMatchIn(nl) || nl.startsWith("Четным пригородным", true) || nl.startsWith("Четные и нечетные", true) || nl.startsWith("Поздний выход", true)) break
                blockLines.add(nl); j++
            }
            i = j - 1; val fullBlock = blockLines.joinToString(" ")
            if (fullBlock.contains("Торжок", true) || fullBlock.contains("Бологое", true) || fullBlock.contains("Угловка", true)) { i++; continue }
            
            val allParens = Regex("\\(([^)]+)\\)").findAll(line).map { it.groupValues[1] }.toList()
            val dateParen = allParens.firstOrNull { Regex("\\d{1,2}[/.\\-]\\d{1,2}|нояб|окт").containsMatchIn(it.lowercase()) }
            val inlineDate = Regex("(?i)(?:только\\s*\\d{1,2}[/.\\-]\\d{1,2}|с\\s*\\d{1,2}[/.\\-]\\d{1,2}\\s*по\\s*\\d{1,2}[/.\\-]\\d{1,2}[^)]*|\\d{1,2}(?:,\\s*\\d{1,2})+/[01]?\\d)").find(line)?.value
            val rawDates = dateParen ?: inlineDate ?: currentSectionDates.ifEmpty { defaultDateGroup }
            val calendarDaySet = parseCalendarDays(rawDates, defaultDateGroup)
            
            val partialMatch = Regex("(?i)\\(на\\s+участке\\s+([^)]+)\\)").find(line)
            val partialCancel = partialMatch?.groupValues?.get(1)?.trim() ?: ""
            val routeClean = Regex("([А-Яа-яЁё0-9\\s.\\-]+?\\s*[-–—]\\s*[А-Яа-яЁё0-9\\s.\\-]+)").find(line.substringAfter(trainNum).replace(":", " ").replace(Regex("(?i)отменен[а-я]*|отменить|назначен[а-я]*"), " ").replace(Regex("\\([^)]*\\)"), " ").replace(Regex("(?i)отправлением.*"), " ").replace(Regex("(?i)только\\s*\\d{1,2}[/.\\-]\\d{1,2}.*"), " ").replace(Regex("(?i)с\\s*\\d{1,2}[/.\\-]\\d{1,2}.*"), " ").trim())?.groupValues?.get(1)?.trim() ?: ""
            
            val skippedList = mutableListOf<String>()
            for (sm in Regex("(?i)\\(([^)]*?)\\s+проследу(?:ет|ют)\\s+без\\s+остановки\\)").findAll(fullBlock)) sm.groupValues[1].split(",").forEach { part -> val norm = normalizeStationKey(part); if (norm.isNotEmpty() && norm !in HARD_CLOSED_STATIONS) { val disp = displayStationName(norm); if (disp !in skippedList) skippedList.add(disp) } }
            
            val appointedList = mutableListOf<String>()
            for (am in Regex("(?i)\\(([^)]*?)\\s+назнача(?:ется|ются)\\s+остановк[аи]\\)").findAll(fullBlock)) {
                am.groupValues[1].split(",").forEach { part -> 
                    val norm = normalizeStationKey(part)
                    if (norm.isNotEmpty()) { 
                        if (norm in HARD_CLOSED_STATIONS) {
                            errors.add(GlobalError(java.util.UUID.randomUUID().toString(), trainNum, telHeader, "Попытка назначить остановку на закрытой станции: ${displayStationName(norm)}"))
                        }
                        val disp = displayStationName(norm)
                        if (disp !in appointedList) appointedList.add(disp) 
                    } 
                }
            }
            
            val trackRules = mutableListOf<TrackSegmentRule>()
            for (trm in Regex("(?i)от\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+до\\s+([А-Яа-яЁё0-9\\s.\\-]+?)\\s+(?:проследует\\s+)?по\\s+(\\d)\\s+(?:неправильному\\s+)?пути").findAll(fullBlock)) {
                val fromKey = normalizeStationKey(trm.groupValues[1]); val toKey = normalizeStationKey(trm.groupValues[2])
                if (fromKey.isNotEmpty() && toKey.isNotEmpty()) trackRules.add(TrackSegmentRule(fromKey, toKey, trm.groupValues[3]))
            }
            
            var explicitNoStops = ""
            val sectionNoStopsMatch = Regex("(?i)(?:на\\s+участке\\s+|от\\s+)?([А-Яа-яЁё\\s\\-]+?[-–—][А-Яа-яЁё\\s\\-]+?)\\s+(?:проследует|следует)\\s+без\\s+остановок").find(fullBlock)
            if (sectionNoStopsMatch != null) explicitNoStops = "${sectionNoStopsMatch.groupValues[1].trim()} проследует без остановок"
            else if (Regex("(?i)(?:проследует|следует)\\s+без\\s+остановок").containsMatchIn(fullBlock)) explicitNoStops = "Следует без остановок"
            
            val bodyWithoutParens = fullBlock.replace(Regex("\\([^)]*\\)"), " "); val stMap = LinkedHashMap<String, Pair<String, Boolean>>()
            val timePointPattern = "\\d{1,2}[:.]\\d{2}(?:[:.]\\d{2})?(?:,\\d+)?"
            val segRegex = Regex("([А-Яа-яЁё][А-Яа-яЁё0-9\\s.\\-]{1,28}?)\\s+(?:приб\\.?\\s*$timePointPattern\\s+подача[^,]+?\\s+)?($timePointPattern)(?:\\s*[-–—:]\\s*($timePointPattern))?")
            val segments = bodyWithoutParens.split(",")
            for ((idx, rawSeg) in segments.withIndex()) {
                val seg = rawSeg.trim(); if (seg.isEmpty()) continue
                val m = segRegex.find(seg) ?: continue
                val normSt = normalizeStationKey(m.groupValues[1].replace(Regex("(?i)действующим\\s+расписанием"), "").replace(Regex("(?i)согласно\\s+ранее.*"), "").replace(Regex("(?i)до\\s+ст\\.?"), "").replace(Regex("(?i)далее.*"), "").trim())
                if (normSt.isEmpty()) continue
                val isTech = Regex("(?i)\\(тех\\.?\\s*стоянка\\)").containsMatchIn(seg) || Regex("(?i)\\(тех\\.?\\s*остановка\\)").containsMatchIn(seg) || seg.contains("тех. стоянка", true)
                if (isTech) continue
                
                val t1 = m.groupValues[2]; val t2 = m.groupValues[3]
                val origSegInFull = fullBlock.substringAfter(m.groupValues[1], "").take(AppConstants.TRACK_HINT_LOOKAHEAD)
                val isTechStop = origSegInFull.contains("тех", true) || seg.contains("без пассажиров", true)
                val depTime = if (t2.isNotEmpty()) cleanSingleTimeToken(t2) else cleanSingleTimeToken(t1)
                val isFirstOrLastSeg = (idx == 0 || idx == segments.lastIndex)
                val hasPassengerStop = (t2.isNotEmpty() || isFirstOrLastSeg || seg.contains("отпр", true) || seg.contains("приб", true)) && !isTechStop
                if (depTime.isNotEmpty()) stMap[normSt] = Pair(depTime, hasPassengerStop)
            }
            val inlineCancel = line.contains("отменен", true) || line.contains("отменить", true)
            events.add(TelegramTrainEvent(trainNum, rawDates, calendarDaySet, telHeader, routeClean, (inCancelBlock || inlineCancel) && partialCancel.isEmpty(), partialCancel, inAppointBlock, stMap, skippedList, appointedList, trackRules, explicitNoStops, fullBlock))
        }
        i++
    }
    return Pair(events, errors)
}

fun applyClarificationTelegrams(events: List<TelegramTrainEvent>): List<TelegramTrainEvent> {
    val clarifications = events.filter { it.isClarificationNotice }
    if (clarifications.isEmpty()) return events

    val standardEvents = events.filter { !it.isClarificationNotice }.toMutableList()

    for (cl in clarifications) {
        val targets = standardEvents.filter { it.trainNum == cl.trainNum }
        if (targets.isEmpty()) {
            standardEvents.add(cl)
            continue
        }

        for (target in targets) {
            val idx = standardEvents.indexOf(target)
            if (idx == -1) continue

            val newAppointed = (target.appointedStops + cl.appointedStops).distinct()
            val newSkipped = target.skippedStops.filter { it !in cl.appointedStops }

            val newStationTimes = LinkedHashMap(target.stationTimes)
            cl.stationTimes.forEach { (st, pair) ->
                val prev = newStationTimes[st]
                if (prev != null) {
                    newStationTimes[st] = Pair(prev.first, pair.second)
                }
            }

            val updatedSkipped = (newSkipped + cl.skippedStops).distinct()
            cl.skippedStops.forEach { st ->
                val normSt = normalizeStationKey(st)
                val prev = newStationTimes[normSt]
                if (prev != null) {
                    newStationTimes[normSt] = Pair(prev.first, false)
                }
            }

            val routeNotice = if (cl.clarificationRouteNotice.isNotEmpty()) cl.clarificationRouteNotice else target.clarificationRouteNotice

            standardEvents[idx] = target.copy(
                appointedStops = newAppointed,
                skippedStops = updatedSkipped,
                stationTimes = newStationTimes,
                clarificationRouteNotice = routeNotice,
                rawBlockText = target.rawBlockText + "\n[Разъяснение " + cl.telegramHeader + "]: " + cl.rawBlockText
            )
        }
    }

    return standardEvents
}

private data class TelegramCacheKey(val paths: List<String>, val lastModified: Long)
object TelegramParseCache {
    private var key: TelegramCacheKey? = null; private var cached: List<TelegramTrainEvent> = emptyList(); private var cachedErrors: List<GlobalError> = emptyList()
    fun get(files: List<File>): Pair<List<TelegramTrainEvent>, List<GlobalError>> {
        if (files.isEmpty()) { key = null; cached = emptyList(); cachedErrors = emptyList(); return Pair(emptyList(), emptyList()) }
        val newKey = TelegramCacheKey(files.map { it.absolutePath }, files.sumOf { it.lastModified() })
        if (newKey == key) return Pair(cached, cachedErrors)
        val results = files.map { parseTelegramDocument(it) }
        val rawEvents = results.flatMap { it.first }
        cached = applyClarificationTelegrams(rawEvents)
        cachedErrors = results.flatMap { it.second }; key = newKey
        return Pair(cached, cachedErrors)
    }
}

private data class OtmenyCacheKey(val paths: List<String>, val lastModified: Long)
object OtmenyParseCache {
    private var key: OtmenyCacheKey? = null; private var cached: Map<String, String> = emptyMap()
    fun get(files: List<File>): Map<String, String> {
        if (files.isEmpty()) { key = null; cached = emptyMap(); return emptyMap() }
        val newKey = OtmenyCacheKey(files.map { it.absolutePath }, files.sumOf { it.lastModified() })
        if (newKey == key) return cached
        cached = parseOtmenyFiles(files); key = newKey; return cached
    }
}

fun parseOtmenyFiles(files: List<File>): Map<String, String> {
    val map = mutableMapOf<String, String>()
    files.forEach { f ->
        readDocumentText(f).lines().map { preprocessTelegramLine(it) }.forEach { line ->
            val m = Regex("№\\s*(\\d{4}(?:/\\d{4})?)\\s+(.*?-\\s*.*)").find(line)
            if (m != null) map[m.groupValues[1]] = m.groupValues[2].trim()
            else {
                val appM = Regex("(?i)Назначается\\s+№\\s*(\\d{4}(?:/\\d{4})?)\\s+(.*)").find(line)
                if (appM != null) map[appM.groupValues[1]] = "Назначается ${appM.groupValues[2].trim()}"
            }
        }
    }
    return map
}

fun parseManualExport(files: List<File>): Map<String, String> {
    val map = mutableMapOf<String, String>()
    files.forEach { f ->
        readDocumentText(f).lines().map { preprocessTelegramLine(it) }.forEach { line ->
            val m = Regex("№\\s*(\\d{4}(?:/\\d{4})?)\\s+(.*?-\\s*.*)").find(line)
            if (m != null) {
                val num = m.groupValues[1]
                val rest = m.groupValues[2].trim()
                val change = rest.substringAfterLast("-").trim()
                map[num] = change
            }
        }
    }
    return map
}

fun buildRemainingStopsNote(targetKey: String, stationTimes: LinkedHashMap<String, Pair<String, Boolean>>, isAppointed: Boolean, explicitNoStopsText: String): String {
    if (explicitNoStopsText.isNotEmpty()) return explicitNoStopsText
    val keys = stationTimes.keys.toList(); val idx = keys.indexOf(targetKey)
    if (idx == -1 || idx >= keys.lastIndex) return ""
    val afterKeys = keys.subList(idx + 1, keys.size)
    val intermediateStops = afterKeys.dropLast(1).filter { k -> (stationTimes[k]?.second ?: false) && k !in HARD_CLOSED_STATIONS }
    if (intermediateStops.isEmpty() && afterKeys.isNotEmpty()) {
        val endStDisplay = displayStationName(afterKeys.last())
        return if (isAppointed) "Следует без остановок." else "До ст. $endStDisplay проследует без остановок."
    } else if (isAppointed && intermediateStops.size in 1..3) {
        return "Следует с остановками: ${(intermediateStops + afterKeys.last()).map { displayStationName(it) }.joinToString(", ")}."
    }
    return ""
}

fun mergeSameTrainRowsForDay(rows: List<StationComparisonRow>): List<StationComparisonRow> {
    val grouped = rows.groupBy { it.trainNum to it.direction }
    val mergedList = mutableListOf<StationComparisonRow>()

    for ((_, rowList) in grouped) {
        if (rowList.size == 1) {
            mergedList.add(rowList.first())
            continue
        }

        val cancelRow = rowList.firstOrNull { it.isCancelledAtStation && !it.isPartialCancel && !it.isPassWithoutStop }
        if (cancelRow != null) {
            mergedList.add(cancelRow)
            continue
        }

        val baseRow = rowList.firstOrNull { it.telegramTime != "-" && it.telegramTime != "ОТМ" } ?: rowList.first()
        val allChanges = mutableListOf<String>()

        val shiftText = rowList.map { it.finalChangeText }.firstOrNull { it.startsWith("отправится", true) }
        val followsText = rowList.map { it.finalChangeText }.firstOrNull { it.contains("следует до", true) || it.contains("посадки нет", true) }
        val passWithoutStop = rowList.any { it.isPassWithoutStop }

        if (passWithoutStop) {
            allChanges.add("ОТМЕНЕН (проследует без остановки)")
        } else {
            if (shiftText != null) {
                val cleanShift = shiftText.substringBefore("следует до").substringBefore(".").trim()
                allChanges.add(cleanShift)
            }
            if (followsText != null) {
                val cleanFollows = if (followsText.contains("следует до")) {
                    "следует до " + followsText.substringAfter("следует до").substringBefore(".").trim()
                } else if (followsText.contains("посадки нет")) {
                    "посадки нет"
                } else ""
                if (cleanFollows.isNotEmpty() && !allChanges.any { it.contains(cleanFollows) }) {
                    allChanges.add(cleanFollows)
                }
            }
        }

        val allSkipped = rowList.flatMap { it.skippedStopsSummary.split(", ") }.filter { it.isNotBlank() }.distinct()
        val allAppointed = rowList.flatMap { it.appointedStopsSummary.split(", ") }.filter { it.isNotBlank() }.distinct()

        val combinedChangeTextBuilder = StringBuilder()
        if (allChanges.isNotEmpty()) {
            combinedChangeTextBuilder.append(allChanges.joinToString(" "))
        }
        if (allSkipped.isNotEmpty()) {
            if (combinedChangeTextBuilder.isNotEmpty()) combinedChangeTextBuilder.append(". ")
            combinedChangeTextBuilder.append("Отменены остановки: ${allSkipped.joinToString(", ")}.")
        }
        if (allAppointed.isNotEmpty()) {
            if (combinedChangeTextBuilder.isNotEmpty()) combinedChangeTextBuilder.append(". ")
            combinedChangeTextBuilder.append("Назначены остановки: ${allAppointed.joinToString(", ")}.")
        }

        val finalString = combinedChangeTextBuilder.toString().ifEmpty { baseRow.finalChangeText }

        mergedList.add(
            baseRow.copy(
                finalChangeText = finalString,
                skippedStopsSummary = allSkipped.joinToString(", "),
                appointedStopsSummary = allAppointed.joinToString(", "),
                rawTelegramText = rowList.map { it.rawTelegramText }.distinct().joinToString("\n---\n")
            )
        )
    }

    return mergedList
}

fun expandAndGroupChronologically(rawItems: List<Pair<CalendarDaySet, StationComparisonRow>>, selectedDates: Set<SingleDateFilterOption>): List<StationComparisonRow> {
    if (rawItems.isEmpty()) return emptyList()
    if (selectedDates.isNotEmpty()) {
        val filtered = rawItems.filter { (cal, _) -> cal.days.isEmpty() || selectedDates.any { it.day in cal.days && it.monthNum == cal.monthNum } }
            .map { (_, row) -> row.copy(dateGroup = selectedDates.joinToString(", ") { it.displayLabel }, sortDayKey = selectedDates.minOfOrNull { it.day } ?: 1, activeDays = selectedDates.map { it.day }, monthName = selectedDates.firstOrNull()?.monthName ?: "") }
        
        return mergeSameTrainRowsForDay(filtered)
            .sortedBy { if (it.afishaTime != "-") it.afishaTime else it.telegramTime }
    }
    val withCalendar = rawItems.filter { it.first.days.isNotEmpty() }
    val withoutCalendar = rawItems.filter { it.first.days.isEmpty() }.map { it.second }
    if (withCalendar.isEmpty()) return withoutCalendar
    val monthName = withCalendar.first().first.monthName
    val allDaysSorted = withCalendar.flatMap { it.first.days }.distinct().sorted()
    val dayToRows = LinkedHashMap<Int, List<StationComparisonRow>>()
    for (day in allDaysSorted) {
        val rowsForDay = withCalendar.filter { day in it.first.days }.map { it.second }
        val mergedForDay = mergeSameTrainRowsForDay(rowsForDay)
            .sortedBy { if (it.afishaTime != "-") it.afishaTime else it.telegramTime }
        if (mergedForDay.isNotEmpty()) dayToRows[day] = mergedForDay
    }
    val result = mutableListOf<StationComparisonRow>()
    var currentDaysGroup = mutableListOf<Int>(); var currentSignature = ""; var currentTemplateRows = emptyList<StationComparisonRow>()
    fun flushGroup() {
        if (currentDaysGroup.isEmpty()) return
        val header = formatDaysBlockHeader(currentDaysGroup, monthName); val firstDay = currentDaysGroup.first(); val daysCopy = currentDaysGroup.toList()
        currentTemplateRows.forEach { r -> result.add(r.copy(dateGroup = header, sortDayKey = firstDay, activeDays = daysCopy, monthName = monthName)) }
    }
    for ((day, rows) in dayToRows) {
        val sig = rows.joinToString("|") { "${it.trainNum}:${it.direction}:${it.finalChangeText}" }
        val isConsecutive = currentDaysGroup.isEmpty() || (day == currentDaysGroup.last() + 1)
        if (currentDaysGroup.isEmpty()) { currentDaysGroup.add(day); currentSignature = sig; currentTemplateRows = rows }
        else if (sig == currentSignature && isConsecutive) { currentDaysGroup.add(day) }
        else { flushGroup(); currentDaysGroup = mutableListOf(day); currentSignature = sig; currentTemplateRows = rows }
    }
    flushGroup(); return result + withoutCalendar
}

private fun makeComparisonRow(targetStationDisplay: String, dirNum: String, ev: TelegramTrainEvent?, afInfo: AfishaTrainInfo?, dir: Direction, isMcd: Boolean, statusText: String, otmStatus: String, isCanc: Boolean, isPassNoStop: Boolean, isApp: Boolean, isShifted: Boolean, isPartial: Boolean, isAppStop: Boolean, trackBadge: String, afTimeAtStation: String, telTimeAtStation: String, calendarDays: CalendarDaySet, telegramHeader: String, rawText: String, validationError: String): StationComparisonRow {
    val hasDiscrep = otmStatus.isNotEmpty() && statusText.isNotEmpty() && !otmStatus.contains(telTimeAtStation.ifEmpty { "___" }) && !(otmStatus.contains("ОТМЕНЕН", true) && isCanc)
    val initialDateHeader = if (calendarDays.days.isNotEmpty()) formatDaysBlockHeader(calendarDays.days, calendarDays.monthName) else calendarDays.fallbackLabel
    return StationComparisonRow(
        dateGroup = initialDateHeader, sortDayKey = calendarDays.days.firstOrNull() ?: AppConstants.UNGROUPED_DAY_KEY, activeDays = calendarDays.days, monthName = calendarDays.monthName,
        trainNum = dirNum, daysFlag = afInfo?.daysFlag ?: "", routeString = afInfo?.routeString ?: ev?.routeFromTelegram?.ifEmpty { null } ?: run { when { dir == Direction.TO_MOSCOW && isMcd -> "Крюково – Ипподром"; dir == Direction.TO_MOSCOW -> "Крюково – Москва"; isMcd -> "Ипподром – Крюково"; else -> "Москва – Крюково" } },
        destinationLabel = afInfo?.destinationLabel ?: when (dir) {
            Direction.TO_MOSCOW -> if (isMcd) "до Ипподрома (47 км)" else "до Москвы"
            Direction.FROM_MOSCOW -> when { afInfo?.routeString?.contains("Твер", true) == true -> "до Твери"; afInfo?.routeString?.contains("Клин", true) == true -> "до Клина"; afInfo?.routeString?.contains("Конаков", true) == true -> "до Конаково ГРЭС"; afInfo?.routeString?.contains("Подсолнеч", true) == true -> "до Подсолнечной"; afInfo?.routeString?.contains("Решетник", true) == true -> "до Решетниково"; else -> "до Крюково" }
        },
        isMcd = isMcd, direction = dir, stationName = targetStationDisplay, telegramHeader = telegramHeader, afishaTime = afTimeAtStation.ifEmpty { "-" }, telegramTime = telTimeAtStation.ifEmpty { if (isCanc) "ОТМ" else "-" },
        otmenyFileStatus = otmStatus, finalChangeText = statusText.ifEmpty { otmStatus }, isCancelledAtStation = isCanc, isPassWithoutStop = isPassNoStop, isAppointedAtStation = isApp, isTimeShiftedAtStation = isShifted, isPartialCancel = isPartial, isAppointedStop = isAppStop,
        skippedStopsSummary = ev?.skippedStops?.joinToString(", ") ?: "", appointedStopsSummary = ev?.appointedStops?.joinToString(", ") ?: "", trackSummaryForStation = trackBadge, hasDiscrepancyWithOtmeny = hasDiscrep,
        fullAfishaStations = afInfo?.stationTimes ?: emptyMap(), fullTelegramStations = ev?.stationTimes ?: emptyMap(), rawTelegramText = rawText, validationError = validationError
    )
}

fun buildStationComparison(targetStationDisplay: String, afishaDb: Map<String, AfishaTrainInfo>, telegramEvents: List<TelegramTrainEvent>, otmenyMap: Map<String, String>, manualExportMap: Map<String, String>, selectedDates: Set<SingleDateFilterOption>, ignoredErrors: Set<String>): Pair<List<StationComparisonRow>, List<StationComparisonRow>> {
    val targetKey = normalizeStationKey(targetStationDisplay); val targetOrder = STATION_ORDER_INDEX[targetKey] ?: AppConstants.DEFAULT_STATION_ORDER
    val toMoscowRaw = mutableListOf<Pair<CalendarDaySet, StationComparisonRow>>(); val fromMoscowRaw = mutableListOf<Pair<CalendarDaySet, StationComparisonRow>>()
    
    if (targetKey in HARD_CLOSED_STATIONS) return Pair(emptyList(), emptyList())
    
    for (ev in telegramEvents) {
        val displayNum = if (ev.trainNum.contains("/")) ev.trainNum else ev.trainNum
        for (dirNum in splitDirectionalTrainNumbers(ev.trainNum)) {
            val afInfo = lookupAfisha(afishaDb, dirNum); val dir = afInfo?.direction ?: getDirectionBySingleNumber(dirNum)
            val isMcd = afInfo?.isMcd ?: ("/" in ev.trainNum || ev.routeFromTelegram.contains("Ипподром", true) || isTrainMcd(dirNum))
            
            if (targetKey == "левобережная" && dir == Direction.FROM_MOSCOW) continue
            
            val afTimeAtStation = afInfo?.stationTimes?.get(targetKey) ?: ""; val telPairAtStation = ev.stationTimes[targetKey]
            val telTimeAtStation = telPairAtStation?.first ?: ""; val telHasStop = telPairAtStation?.second ?: false
            val isStationInSkippedList = ev.skippedStops.any { normalizeStationKey(it) == targetKey }; val isStationInAppointedList = ev.appointedStops.any { normalizeStationKey(it) == targetKey }
            val trackBadge = resolveTrackForStation(targetKey, dirNum, dir, ev.trackRules, afInfo?.stationTrackHints?.get(targetKey) ?: "")
            val trackSuffix = if (trackBadge.isNotEmpty()) " $trackBadge" else ""
            val passesStation = afTimeAtStation.isNotEmpty() || telPairAtStation != null || isStationInSkippedList || isStationInAppointedList || (afishaDb.isEmpty() && (ev.isCancelledFull || ev.isAppointedFull))
            if (!passesStation) continue
            
            var validationError = ""
            if (targetKey in HARD_CLOSED_STATIONS && isStationInAppointedList) {
                validationError = "Остановка на закрытой станции"
            }

            var statusText = ""; var isCanc = false; var isPassNoStop = false; var isApp = false; var isShifted = false; var isPartial = false; var isAppStop = false
            if (ev.isCancelledFull) { 
                isCanc = true; statusText = "ОТМЕНЕН" 
            }
            else if (ev.partialCancelSection.isNotEmpty()) {
                val secNorm = ev.partialCancelSection.lowercase()
                if (("крюков" in secNorm || "зеленогр" in secNorm) && "подсолнеч" in secNorm && dir == Direction.FROM_MOSCOW) { 
                    statusText = if (telTimeAtStation.isNotEmpty() && afTimeAtStation.isNotEmpty() && telTimeAtStation != afTimeAtStation) { 
                        isShifted = true
                        "отправится в $telTimeAtStation$trackSuffix следует до Крюково" 
                    } else "следует до Крюково" 
                } else if ("клин" in secNorm && "решетник" in secNorm && dir == Direction.FROM_MOSCOW) statusText = "следует до Клина"
                else if ("поваров" in secNorm && "клин" in secNorm && dir == Direction.FROM_MOSCOW) statusText = "следует до Поварово-1"
                else { isCanc = true; isPartial = true; statusText = "ОТМЕНЕН (на уч. ${ev.partialCancelSection})" }
            } else if (ev.isAppointedFull) {
                if (telTimeAtStation.isEmpty() && afTimeAtStation.isEmpty()) continue
                isApp = true; val t = telTimeAtStation.ifEmpty { afTimeAtStation }; val stopsNote = buildRemainingStopsNote(targetKey, ev.stationTimes, true, ev.explicitNoStopsText)
                statusText = "НАЗНАЧЕН в $t$trackSuffix" + if (stopsNote.isNotEmpty()) ". $stopsNote" else ""
            } else {
                val timeActuallyChanged = afTimeAtStation.isNotEmpty() && telTimeAtStation.isNotEmpty() && telTimeAtStation != afTimeAtStation
                if (isStationInAppointedList) { isAppStop = true; statusText = "Назначена остановка ${displayStationName(targetKey)}, отправится в ${telTimeAtStation.ifEmpty { afTimeAtStation }}$trackSuffix" }
                else if (isStationInSkippedList || (telPairAtStation != null && !telHasStop && afTimeAtStation.isNotEmpty())) { isCanc = true; isPassNoStop = true; statusText = "ОТМЕНЕН (проследует без остановки)" }
                else if (telTimeAtStation.isNotEmpty() && telHasStop) { if (timeActuallyChanged || trackBadge.isNotEmpty() || afishaDb.isEmpty()) { isShifted = timeActuallyChanged || afishaDb.isEmpty(); statusText = "отправится в $telTimeAtStation$trackSuffix" } }
                
                if (ev.clarificationRouteNotice.isNotEmpty()) {
                    val notice = ev.clarificationRouteNotice
                    statusText = if (statusText.isNotEmpty()) "$statusText $notice" else notice
                }

                if (!isCanc) {
                    val otherSkippedAhead = ev.skippedStops.filter { k -> val ord = STATION_ORDER_INDEX[normalizeStationKey(k)] ?: return@filter false; if (dir == Direction.TO_MOSCOW) ord > targetOrder else ord < targetOrder }
                    val otherAppointedAhead = ev.appointedStops.filter { k -> val ord = STATION_ORDER_INDEX[normalizeStationKey(k)] ?: return@filter false; if (dir == Direction.TO_MOSCOW) ord > targetOrder else ord < targetOrder }
                    if (otherSkippedAhead.isNotEmpty()) statusText = (if (statusText.isNotEmpty()) "$statusText. " else "") + "Отменены остановки: ${otherSkippedAhead.joinToString(", ")}."
                    if (otherAppointedAhead.isNotEmpty()) statusText = (if (statusText.isNotEmpty()) "$statusText. " else "") + "Назначены остановки: ${otherAppointedAhead.joinToString(", ")}."
                    if (otherSkippedAhead.isEmpty() && ev.explicitNoStopsText.isNotEmpty()) statusText = (if (statusText.isNotEmpty()) "$statusText. " else "") + ev.explicitNoStopsText
                }
            }
            
            var otmStatus = lookupOtmenyStatus(otmenyMap, dirNum)
            if (manualExportMap.isNotEmpty()) {
                val manVal = manualExportMap[displayNum] ?: manualExportMap[dirNum]
                if (manVal != null) {
                    val expected = statusText.ifEmpty { "Без изменений" }
                    if (!manVal.contains(expected, true) && !expected.contains(manVal, true)) {
                        otmStatus = (if (otmStatus.isNotEmpty()) "$otmStatus. " else "") + "РУЧНАЯ ВЫГРУЗКА: $manVal"
                    }
                } else {
                    otmStatus = (if (otmStatus.isNotEmpty()) "$otmStatus. " else "") + "НЕТ В РУЧНОЙ ВЫГРУЗКЕ!"
                }
            }

            if (statusText.isEmpty() && otmStatus.isEmpty() && validationError.isEmpty()) continue
            if (validationError.isNotEmpty() && ignoredErrors.contains("${ev.trainNum}_${targetKey}")) { validationError = "" }
            
            val row = makeComparisonRow(targetStationDisplay, if (isMcd) displayNum else dirNum, ev, afInfo, dir, isMcd, statusText, otmStatus, isCanc, isPassNoStop, isApp, isShifted, isPartial, isAppStop, trackBadge, afTimeAtStation, telTimeAtStation, ev.calendarDays, ev.telegramHeader, ev.rawBlockText, validationError)
            
            if (dir == Direction.TO_MOSCOW) toMoscowRaw.add(Pair(ev.calendarDays, row)) else fromMoscowRaw.add(Pair(ev.calendarDays, row))
        }
    }
    for ((num, otmStat) in otmenyMap) {
        for (dirNum in splitDirectionalTrainNumbers(num)) {
            val afInfo = lookupAfisha(afishaDb, dirNum); val afTime = afInfo?.stationTimes?.get(targetKey) ?: "-"; val dir = afInfo?.direction ?: getDirectionBySingleNumber(dirNum)
            
            if (targetKey == "левобережная" && dir == Direction.FROM_MOSCOW) continue
            
            val isMcd = afInfo?.isMcd ?: isTrainMcd(dirNum); val emptyCal = CalendarDaySet(AppConstants.DEFAULT_MONTH_NUM, AppConstants.DEFAULT_MONTH_NAME, emptyList(), "По файлу «Отмены»")
            val row = makeComparisonRow(targetStationDisplay, if (isMcd) num else dirNum, null, afInfo, dir, isMcd, "", otmStat, otmStat.contains("ОТМЕНЕН", true), otmStat.contains("без остановки", true), otmStat.contains("Назнач", true), otmStat.contains("отправится", true), false, false, "", afTime, "", emptyCal, "Нет в телеграмме!", "Поезд присутствует в файле выгрузки «Отмены», но не найден в загруженных телеграммах.", "").copy(hasDiscrepancyWithOtmeny = true)
            if (dir == Direction.TO_MOSCOW) toMoscowRaw.add(Pair(emptyCal, row)) else fromMoscowRaw.add(Pair(emptyCal, row))
        }
    }
    return Pair(expandAndGroupChronologically(toMoscowRaw, selectedDates), expandAndGroupChronologically(fromMoscowRaw, selectedDates))
}

private data class StationCalcResult(val toMoscow: List<StationComparisonRow>, val fromMoscow: List<StationComparisonRow>, val availableDates: List<SingleDateFilterOption>, val telegramEvents: List<TelegramTrainEvent>, val otmenyMap: Map<String, String>, val errors: List<GlobalError>)

fun main() = application {
    Window(onCloseRequest = ::exitApplication, title = "Train-R — Интеллектуальный ревизор пригородных расписаний", state = WindowState(width = 1360.dp, height = 860.dp)) {
        TrainRApp()
    }
}
