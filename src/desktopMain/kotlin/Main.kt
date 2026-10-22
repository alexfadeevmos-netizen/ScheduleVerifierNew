import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.Canvas
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.horizontalScroll
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.LazyRow
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
import androidx.compose.ui.text.font.FontStyle
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
import java.util.UUID

// ==================== 1. КОНСТАНТЫ И ПЕРЕЧИСЛЕНИЯ ====================

object AppConstants {
    const val DEFAULT_MONTH_NUM = 10
    const val DEFAULT_MONTH_NAME = "октября"
    const val UNGROUPED_DAY_KEY = 999
    const val DEFAULT_STATION_ORDER = 27
    const val TRACK_HINT_LOOKAHEAD = 45
    const val REPARSE_DEBOUNCE_MS = 150L
    const val SEARCH_DEBOUNCE_MS = 350L
}

enum class Direction { TO_MOSCOW, FROM_MOSCOW }
enum class TrainFilter { ALL, MCD, EXPRESS }
enum class YandexDiscrepancySeverity { CRITICAL_CANCEL_MISSED, TIME_MISMATCH, STOP_MISMATCH, INFO }

// ==================== 2. МОДЕЛИ ДАННЫХ ====================

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
    val isCancelledAtStation: Boolean, val isPassWithoutStop: Boolean, val isAppointedAtStation: Boolean, 
    val isTimeShiftedAtStation: Boolean, // ВОЗВРАЩЕНО СВОЙСТВО
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
    val rawText: String = ""
)

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

private data class StationCalcResult(
    val toMoscow: List<StationComparisonRow>,
    val fromMoscow: List<StationComparisonRow>,
    val availableDates: List<SingleDateFilterOption>,
    val telegramEvents: List<TelegramTrainEvent>,
    val otmenyMap: Map<String, String>,
    val errors: List<GlobalError>
)

// ==================== 3. ТЕМА ОФОРМЛЕНИЯ И СПРАВОЧНИКИ ====================

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

// ==================== 4. БАЗОВЫЕ УТИЛИТЫ ====================

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
                FileInputStream(file).use { fis -> XWPFDocument(fis).use { doc -> doc.paragraphs.joinToString("\n") { it.text } } }
            }
            file.name.endsWith(".pdf", true) -> {
                Loader.loadPDF(file).use { doc -> PDFTextStripper().getText(doc) }
            }
            else -> file.readText(Charsets.UTF_8)
        }
    } catch (e: Exception) { logWarn("Ошибка чтения файла ${file.name}", e); "" }
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
        if (inCancelBlock || inAppointBlock || rest.contains("отмен", true) || rest.contains("назнач", true) || Regex("[-–—]").containsMatchIn(rest.take(35))) return Pair(true, bareMatch.groupValues[1])
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
        "подсолн" in clean || "подснечн" in clean || "поделонечн" in clean -> "подсолнечная"
        "сенеж" in clean -> "сенеж"
        "головк" in clean -> "головково"
        "покровк" in clean -> "покровка"
        "фролов" in clean -> "фроловское"
        "стрегл" in clean || "стригл" in clean || "стрелк" in clean || "стрелгов" in clean -> "стреглово"
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

// ==================== 5. ПАРСЕРЫ (АФИША, ТЕЛЕГРАММЫ, ЯНДЕКС) ====================

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
                    if (cellVal.isNotEmpty() && !cellVal.contains("Условные", true) && !cellVal.contains("Расписание", true) && !cellVal.equals("изм", true) && !cellVal.equals("отм", true)) stationRows[r] = cellVal                }
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
                            if (hints == null) { hints = mutableMapOf(); trackHintsDb[trainNum] = hints }
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

// [YANDEX] Парсер Excel от Яндекса
fun parseYandexExcel(file: File): List<YandexParsedTrain> {
    val result = mutableListOf<YandexParsedTrain>()
    try {
        FileInputStream(file).use { fis ->
            val wb = WorkbookFactory.create(fis)
            val sheet = wb.getSheetAt(0)
            var headerRowIdx = -1
            for (r in 0..10) {
                val row = sheet.getRow(r) ?: continue
                val rowStr = row.toString().lowercase()
                if (rowStr.contains("время") || rowStr.contains("номер")) { headerRowIdx = r; break }
            }
            if (headerRowIdx == -1) headerRowIdx = 0

            for (r in (headerRowIdx + 1)..sheet.lastRowNum) {
                val row = sheet.getRow(r) ?: continue
                val timeCell = row.getCell(0)?.toString()?.trim() ?: ""
                val routeCell = row.getCell(1)?.toString()?.trim() ?: ""
                val numCell = row.getCell(2)?.toString()?.trim() ?: ""
                val stopsCell = row.getCell(3)?.toString()?.trim() ?: ""

                if (timeCell.matches(Regex("\\d{1,2}:\\d{2}")) && numCell.matches(Regex("\\d{4}(?:/\\d{4})?"))) {
                    result.add(YandexParsedTrain(numCell, timeCell, routeCell, stopsCell))
                }
            }
        }
    } catch (e: Exception) { logWarn("Ошибка парсинга Excel Яндекса", e) }
    return result
}

fun parseTelegramDocument(file: File): Pair<List<TelegramTrainEvent>, List<GlobalError>> {
    val rawText = readDocumentText(file)
    if (rawText.isBlank()) return Pair(emptyList(), emptyList())

    val headerMatch = Regex("(?i)ТЕЛЕГРАММА\\s*№?\\s*([^\\s]+)\\s+от\\s+([0-9/.\\-]+(?:г\\.?)?)").find(rawText)
    val telHeader = if (headerMatch != null) "Тел. №${headerMatch.groupValues[1]} от ${headerMatch.groupValues[2]}" else "Тел. (${file.nameWithoutExtension.take(18)})"
    val globalDatesMatch = Regex("(?i)(?:с\\s*\\d{1,2}\\s+[а-я]+\\s+по\\s+\\d{1,2}\\s+[а-я]+(?:,\\s*с\\s*\\d{1,2}\\s+[а-я]+\\s+по\\s*\\d{1,2}\\s+[а-я]+)?)").find(rawText)
    val defaultDateGroup = globalDatesMatch?.value?.trim() ?: "По датам телеграммы"

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
                        if (norm in HARD_CLOSED_STATIONS) errors.add(GlobalError(UUID.randomUUID().toString(), trainNum, telHeader, "Попытка назначить остановку на закрытой станции: ${displayStationName(norm)}"))
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

private data class TelegramCacheKey(val paths: List<String>, val lastModified: Long)
object TelegramParseCache {
    private var key: TelegramCacheKey? = null; private var cached: List<TelegramTrainEvent> = emptyList(); private var cachedErrors: List<GlobalError> = emptyList()
    fun get(files: List<File>): Pair<List<TelegramTrainEvent>, List<GlobalError>> {
        if (files.isEmpty()) { key = null; cached = emptyList(); cachedErrors = emptyList(); return Pair(emptyList(), emptyList()) }
        val newKey = TelegramCacheKey(files.map { it.absolutePath }, files.sumOf { it.lastModified() })
        if (newKey == key) return Pair(cached, cachedErrors)
        val results = files.map { parseTelegramDocument(it) }
        val rawEvents = results.flatMap { it.first }
        cached = rawEvents
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

// ==================== 6. СВЕРКА ====================

// ВОЗВРАЩЕНА ФУНКЦИЯ buildRemainingStopsNote
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
        if (rowList.size == 1) { mergedList.add(rowList.first()); continue }

        val cancelRow = rowList.firstOrNull { it.isCancelledAtStation && !it.isPartialCancel && !it.isPassWithoutStop }
        if (cancelRow != null) { mergedList.add(cancelRow); continue }

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
                val cleanFollows = if (followsText.contains("следует до")) "следует до " + followsText.substringAfter("следует до").substringBefore(".").trim()
                    else if (followsText.contains("посадки нет")) "посадки нет" else ""
                if (cleanFollows.isNotEmpty() && !allChanges.any { it.contains(cleanFollows) }) allChanges.add(cleanFollows)
            }
        }

        val allSkipped = rowList.flatMap { it.skippedStopsSummary.split(", ") }.filter { it.isNotBlank() }.distinct()
        val allAppointed = rowList.flatMap { it.appointedStopsSummary.split(", ") }.filter { it.isNotBlank() }.distinct()

        val combinedChangeTextBuilder = StringBuilder()
        if (allChanges.isNotEmpty()) combinedChangeTextBuilder.append(allChanges.joinToString(" "))
        if (allSkipped.isNotEmpty()) {
            if (combinedChangeTextBuilder.isNotEmpty()) combinedChangeTextBuilder.append(". ")
            combinedChangeTextBuilder.append("Отменены остановки: ${allSkipped.joinToString(", ")}.")
        }
        if (allAppointed.isNotEmpty()) {
            if (combinedChangeTextBuilder.isNotEmpty()) combinedChangeTextBuilder.append(". ")
            combinedChangeTextBuilder.append("Назначены остановки: ${allAppointed.joinToString(", ")}.")
        }

        val finalString = combinedChangeTextBuilder.toString().ifEmpty { baseRow.finalChangeText }

        mergedList.add(baseRow.copy(finalChangeText = finalString, skippedStopsSummary = allSkipped.joinToString(", "), appointedStopsSummary = allAppointed.joinToString(", "), rawTelegramText = rowList.map { it.rawTelegramText }.distinct().joinToString("\n---\n")))
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
        otmenyFileStatus = otmStatus, finalChangeText = statusText.ifEmpty { otmStatus }, isCancelledAtStation = isCanc, isPassWithoutStop = isPassNoStop, isAppointedAtStation = isApp, 
        isTimeShiftedAtStation = isShifted, // ПЕРЕДАЕТСЯ СВОЙСТВО
        isPartialCancel = isPartial, isAppointedStop = isAppStop,
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
            if (targetKey in HARD_CLOSED_STATIONS && isStationInAppointedList) validationError = "Остановка на закрытой станции"

            var statusText = ""; var isCanc = false; var isPassNoStop = false; var isApp = false; var isShifted = false; var isPartial = false; var isAppStop = false
            if (ev.isCancelledFull) { isCanc = true; statusText = "ОТМЕНЕН" }
            else if (ev.partialCancelSection.isNotEmpty()) {
                val secNorm = ev.partialCancelSection.lowercase()
                if (("крюков" in secNorm || "зеленогр" in secNorm) && "подсолнеч" in secNorm && dir == Direction.FROM_MOSCOW) {
                    statusText = if (telTimeAtStation.isNotEmpty() && afTimeAtStation.isNotEmpty() && telTimeAtStation != afTimeAtStation) { isShifted = true; "отправится в $telTimeAtStation$trackSuffix следует до Крюково" } else "следует до Крюково"
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
                if (ev.clarificationRouteNotice.isNotEmpty()) { val notice = ev.clarificationRouteNotice; statusText = if (statusText.isNotEmpty()) "$statusText $notice" else notice }
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
                    if (!manVal.contains(expected, true) && !expected.contains(manVal, true)) otmStatus = (if (otmStatus.isNotEmpty()) "$otmStatus. " else "") + "РУЧНАЯ ВЫГРУЗКА: $manVal"
                } else otmStatus = (if (otmStatus.isNotEmpty()) "$otmStatus. " else "") + "НЕТ В РУЧНОЙ ВЫГРУЗКЕ!"
            }

            if (statusText.isEmpty() && otmStatus.isEmpty() && validationError.isEmpty()) continue
            if (validationError.isNotEmpty() && ignoredErrors.contains("${ev.trainNum}_${targetKey}")) validationError = ""
            
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

// [YANDEX] Сверка Яндекса с Телеграммами
fun buildYandexDiscrepancies(yandexTrains: List<YandexParsedTrain>, telegramEvents: List<TelegramTrainEvent>, afishaDb: Map<String, AfishaTrainInfo>): List<YandexDiscrepancyRow> {
    if (yandexTrains.isEmpty()) return emptyList()
    val result = mutableListOf<YandexDiscrepancyRow>()

    for (y in yandexTrains) {
        for (dirNum in splitDirectionalTrainNumbers(y.trainNum)) {
            val afisha = lookupAfisha(afishaDb, dirNum)
            val dir = afisha?.direction ?: getDirectionBySingleNumber(dirNum)

            val ev = telegramEvents.firstOrNull { evt -> dirNum in splitDirectionalTrainNumbers(evt.trainNum) }
            val yandexTime = y.departureTime
            val afishaTime = afisha?.stationTimes?.values?.firstOrNull() ?: ""

            if (ev?.isCancelledFull == true) {
                result.add(YandexDiscrepancyRow(dirNum, y.route, dir, yandexTime, "ОТМЕНЕН", y.stopsRule, "Полная отмена", ev.telegramHeader, "Поезд отменён в телеграмме, но есть в Яндексе", YandexDiscrepancySeverity.CRITICAL_CANCEL_MISSED))
                continue
            }

            if (afishaTime.isNotEmpty() && afishaTime != yandexTime && ev?.isTimeShiftedAtStation == true) {
                result.add(YandexDiscrepancyRow(dirNum, y.route, dir, yandexTime, afishaTime, y.stopsRule, "Время изменено", ev.telegramHeader, "Время в Яндексе не совпадает с Афишей", YandexDiscrepancySeverity.TIME_MISMATCH))
            }
        }
    }
    return result.distinctBy { "${it.trainNum}_${it.yandexTime}_${it.severity}" }
}

// ==================== 7. ЭКСПОРТ ====================

fun exportStationFilesToFolders(station: String, toMoscow: List<StationComparisonRow>, fromMoscow: List<StationComparisonRow>, dateFilterSuffix: String = "") {
    val desktop = File(System.getProperty("user.home"), "Desktop")
    val dirToMoscow = File(desktop, "На Москву и Ипподром (нечетные)").apply { if (!exists()) mkdirs() }
    val dirFromMoscow = File(desktop, "На область (из Ипподрома) (четные)").apply { if (!exists()) mkdirs() }
    val fileDatePart = if (dateFilterSuffix.isNotEmpty()) " ($dateFilterSuffix)" else ""
    
    if (toMoscow.isNotEmpty()) {
        try {
            File(dirToMoscow, "$station на Москву$fileDatePart.txt").printWriter(Charsets.UTF_8).use { out ->
                out.println("$station    Изменения в графике движения эл/поездов")
                out.println("на Ленинградский вокзал, Ипподром")
                toMoscow.filter { it.validationError.isEmpty() && it.stationName !in HARD_CLOSED_STATIONS }.groupBy { it.dateGroup }.forEach { (dKey, rowsInGroup) ->
                    out.println(dKey)
                    rowsInGroup.forEach { r ->
                        val timeStr = if (r.afishaTime != "-") "отправлением в ${r.afishaTime}" else ""
                        val change = if (r.isCancelledAtStation && !r.isPartialCancel && !r.isPassWithoutStop) "ОТМЕНЕН" else r.finalChangeText
                        out.println("        №${r.trainNum} ${r.destinationLabel} $timeStr - $change")
                    }
                }
            }
        } catch (e: Exception) { logWarn("Ошибка сохранения", e) }
    }
    if (fromMoscow.isNotEmpty()) {
        try {
            File(dirFromMoscow, "$station от Москвы$fileDatePart.txt").printWriter(Charsets.UTF_8).use { out ->
                out.println("$station   Изменения в графике движения эл/поездов")
                out.println("от Ленинградского вокзала, от Ипподрома")
                fromMoscow.filter { it.validationError.isEmpty() && it.stationName !in HARD_CLOSED_STATIONS }.groupBy { it.dateGroup }.forEach { (dKey, rowsInGroup) ->
                    out.println(dKey)
                    rowsInGroup.forEach { r ->
                        val timeStr = if (r.afishaTime != "-") "отправлением в ${r.afishaTime}" else ""
                        val change = if (r.isCancelledAtStation && !r.isPartialCancel && !r.isPassWithoutStop) "ОТМЕНЕН" else r.finalChangeText
                        out.println("        №${r.trainNum} ${r.destinationLabel} $timeStr - $change")
                    }
                }
            }
        } catch (e: Exception) { logWarn("Ошибка сохранения", e) }
    }
}

// ==================== 8. UI КОМПОНЕНТЫ ====================

@Composable
fun VintageTrainLogo(modifier: Modifier = Modifier.size(46.dp)) {
    Canvas(modifier = modifier) {
        val w = size.width; val h = size.height
        drawCircle(color = Color(0xFF0F172A), radius = w / 2f); drawCircle(color = Color(0xFFD97706), radius = w / 2f - 2f, style = Stroke(width = 3f))
        drawCircle(color = Color(0xFFE2E8F0), radius = w * 0.08f, center = Offset(w * 0.42f, h * 0.25f)); drawCircle(color = Color(0xFFF8FAFC), radius = w * 0.10f, center = Offset(w * 0.55f, h * 0.22f)); drawCircle(color = Color(0xFFE2E8F0), radius = w * 0.07f, center = Offset(w * 0.66f, h * 0.28f))
        drawRect(color = Color(0xFFB91C1C), topLeft = Offset(w * 0.22f, h * 0.42f), size = Size(w * 0.20f, h * 0.26f)); drawRect(color = Color(0xFFFEF3C7), topLeft = Offset(w * 0.26f, h * 0.47f), size = Size(w * 0.11f, h * 0.11f)); drawRect(color = Color(0xFF1E293B), topLeft = Offset(w * 0.42f, h * 0.48f), size = Size(w * 0.30f, h * 0.20f))
        drawRect(color = Color(0xFFD97706), topLeft = Offset(w * 0.60f, h * 0.34f), size = Size(w * 0.08f, h * 0.14f))
        val cowPath = Path().apply { moveTo(w * 0.72f, h * 0.56f); lineTo(w * 0.82f, h * 0.68f); lineTo(w * 0.72f, h * 0.68f); close() }; drawPath(cowPath, color = Color(0xFFD97706))
        drawCircle(color = Color(0xFFD97706), radius = w * 0.08f, center = Offset(w * 0.32f, h * 0.72f)); drawCircle(color = Color(0xFFD97706), radius = w * 0.08f, center = Offset(w * 0.50f, h * 0.72f)); drawCircle(color = Color(0xFFD97706), radius = w * 0.06f, center = Offset(w * 0.66f, h * 0.74f))
        drawLine(color = Color(0xFFF59E0B), start = Offset(w * 0.16f, h * 0.81f), end = Offset(w * 0.84f, h * 0.81f), strokeWidth = 2.5f)
    }
}

@Composable
fun FilterChipItem(label: String, isSelected: Boolean, onClick: () -> Unit) {
    Surface(modifier = Modifier.height(26.dp).clickable { onClick() }, shape = RoundedCornerShape(13.dp), color = if (isSelected) TrainRTheme.RoyalBlue else Color(0xFFF1F5F9), border = BorderStroke(1.dp, if (isSelected) TrainRTheme.RoyalBlue else TrainRTheme.BorderCol)) {
        Box(contentAlignment = Alignment.Center, modifier = Modifier.padding(horizontal = 10.dp)) { Text(label, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = if (isSelected) Color.White else TrainRTheme.TextDark, maxLines = 1) }
    }
}

@Composable
fun CompactFileBadge(title: String, subtitle: String, isLoaded: Boolean, accent: Color, modifier: Modifier, onSelect: () -> Unit, onClear: () -> Unit) {
    Surface(modifier = modifier.height(50.dp).clickable { onSelect() }, shape = RoundedCornerShape(8.dp), color = if (isLoaded) accent.copy(alpha = 0.08f) else Color(0xFFF8FAFC), border = BorderStroke(1.dp, if (isLoaded) accent else TrainRTheme.BorderCol)) {
        Row(modifier = Modifier.fillMaxSize().padding(horizontal = 8.dp, vertical = 4.dp), verticalAlignment = Alignment.CenterVertically) {
            Icon(imageVector = if (isLoaded) Icons.Default.CheckCircle else Icons.Default.Add, contentDescription = null, tint = if (isLoaded) accent else TrainRTheme.TextSub, modifier = Modifier.size(16.dp))
            Spacer(modifier = Modifier.width(6.dp))
            Column(modifier = Modifier.weight(1f), verticalArrangement = Arrangement.Center) {
                Text(title, fontSize = 10.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark, maxLines = 1, overflow = TextOverflow.Ellipsis)
                Text(subtitle, fontSize = 9.sp, fontWeight = FontWeight.SemiBold, color = TrainRTheme.TextSub, maxLines = 1, overflow = TextOverflow.Ellipsis)
            }
            if (isLoaded) IconButton(onClick = onClear, modifier = Modifier.size(22.dp)) { Icon(Icons.Default.Clear, "Удалить", tint = TrainRTheme.RedCancel, modifier = Modifier.size(14.dp)) }
        }
    }
}

@Composable
fun StationChipItem(name: String, isChecked: Boolean, isSelected: Boolean, modifier: Modifier, onCheck: () -> Unit, onSelect: () -> Unit) {
    Surface(modifier = modifier.height(28.dp).clickable { onSelect() }, shape = RoundedCornerShape(6.dp), color = if (isSelected) TrainRTheme.RoyalBlue else Color(0xFFF8FAFC), border = BorderStroke(1.dp, if (isSelected) TrainRTheme.RoyalBlue else TrainRTheme.BorderCol)) {
        Row(modifier = Modifier.fillMaxSize().padding(horizontal = 4.dp), verticalAlignment = Alignment.CenterVertically) {
            Checkbox(checked = isChecked, onCheckedChange = { onCheck() }, modifier = Modifier.size(16.dp), colors = CheckboxDefaults.colors(checkedColor = if (isSelected) Color.White else TrainRTheme.RoyalBlue, checkmarkColor = if (isSelected) TrainRTheme.RoyalBlue else Color.White))
            Spacer(modifier = Modifier.width(4.dp))
            Text(text = name.replace("Зеленоград-", "").replace(" ГРЭС", ""), fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, color = if (isSelected) Color.White else TrainRTheme.NavyDark, maxLines = 1, overflow = TextOverflow.Ellipsis)
        }
    }
}

@Composable
fun StatSummaryBox(label: String, count: Int, color: Color, isSelected: Boolean, modifier: Modifier, onClick: () -> Unit) {
    Surface(modifier = modifier.clickable { onClick() }, shape = RoundedCornerShape(8.dp), color = if (isSelected) color.copy(alpha = 0.25f) else color.copy(alpha = 0.08f), border = BorderStroke(if (isSelected) 2.5.dp else 1.dp, if (isSelected) color else color.copy(alpha = 0.35f))) {
        Column(modifier = Modifier.padding(12.dp)) {
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                Text(label, fontSize = 12.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                Text(count.toString(), fontSize = 22.sp, fontWeight = FontWeight.Black, color = color)
            }
            Text(if (isSelected) "Фильтр активен (нажмите сброс)" else "Нажмите для фильтра", fontSize = 9.sp, fontWeight = FontWeight.Bold, color = color, modifier = Modifier.padding(top = 2.dp))
        }
    }
}

@Composable
fun DetailedReportTrainCard(r: StationComparisonRow, accent: Color, onClick: () -> Unit) {
    val diffText = if (!r.isCancelledAtStation && r.afishaTime != "-" && r.telegramTime != "-" && r.telegramTime != "ОТМ") calculateMinutesDiffLabel(r.afishaTime, r.telegramTime) else ""
    val timeComparisonStr = when { r.isCancelledAtStation -> "${r.afishaTime} ➔ ОТМЕНЕН"; r.isAppointedAtStation && r.afishaTime == "-" -> "Назначен ➔ ${r.telegramTime}"; else -> "${r.afishaTime} ➔ ${r.telegramTime} $diffText".trim() }
    Row(modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF8FAFC)).border(1.dp, accent.copy(alpha = 0.35f), RoundedCornerShape(6.dp)).clickable { onClick() }.padding(vertical = 7.dp, horizontal = 10.dp), verticalAlignment = Alignment.CenterVertically) {
        Text(r.dateGroup, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.RoyalBlue, modifier = Modifier.width(120.dp))
        Text("№${r.trainNum}", fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Black, fontSize = 13.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(95.dp))
        Column(modifier = Modifier.width(175.dp)) {
            Text(r.routeString, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark, maxLines = 1, overflow = TextOverflow.Ellipsis)
            Text(r.destinationLabel, fontWeight = FontWeight.SemiBold, fontSize = 9.sp, color = TrainRTheme.TextSub)
        }
        Surface(shape = RoundedCornerShape(5.dp), color = accent.copy(alpha = 0.12f), modifier = Modifier.width(185.dp).padding(end = 10.dp)) {
            Text(text = timeComparisonStr, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Black, fontSize = 11.sp, color = accent, modifier = Modifier.padding(horizontal = 6.dp, vertical = 3.dp))
        }
        Text(text = r.finalChangeText, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
    }
}

@Composable
fun TrainTableSection(rows: List<StationComparisonRow>, emptyMessage: String, showStationColumn: Boolean = false, onRowClick: (StationComparisonRow) -> Unit) {
    Column(modifier = Modifier.fillMaxSize().padding(6.dp)) {
        Row(modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 6.dp, horizontal = 6.dp), verticalAlignment = Alignment.CenterVertically) {
            Text("ДАТЫ", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(120.dp))
            if (showStationColumn) Text("СТАНЦИЯ", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(95.dp))
            Text("№ ПОЕЗДА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(95.dp))
            Text("МАРШРУТ", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(135.dp))
            Text("ТЕЛЕГРАММА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(115.dp))
            Text("АФИША", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(55.dp))
            Text("ТЕЛЕГР.", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(55.dp))
            Text("ИЗМЕНЕНИЯ", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
            Text("СВЕРКА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(120.dp))
        }
        Spacer(modifier = Modifier.height(2.dp))
        if (rows.isEmpty()) { Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) { Text(emptyMessage, fontSize = 12.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub) } }
        else {
            LazyColumn(modifier = Modifier.fillMaxSize()) {
                items(rows, key = { "${it.trainNum}_${it.direction}_${it.dateGroup}_${it.finalChangeText}" }) { item ->
                    val bg = when { item.validationError.isNotEmpty() -> Color(0xFFFECDD3); item.hasDiscrepancyWithOtmeny -> Color(0xFFFEF3C7); item.isCancelledAtStation -> Color(0xFFFFF1F2); item.isAppointedAtStation || item.isAppointedStop -> Color(0xFFECFDF5); else -> Color.White }
                    val borderCol = when { item.validationError.isNotEmpty() -> TrainRTheme.RedCancel; item.hasDiscrepancyWithOtmeny -> Color(0xFFF59E0B); item.isCancelledAtStation -> Color(0xFFFECDD3); item.isAppointedAtStation || item.isAppointedStop -> Color(0xFFA7F3D0); else -> Color(0xFFE2E8F0) }
                    Row(modifier = Modifier.fillMaxWidth().padding(vertical = 1.dp).clip(RoundedCornerShape(4.dp)).background(bg).border(1.dp, borderCol, RoundedCornerShape(4.dp)).clickable { onRowClick(item) }.padding(vertical = 4.dp, horizontal = 6.dp), verticalAlignment = Alignment.CenterVertically) {
                        Text(item.dateGroup, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.RoyalBlue, modifier = Modifier.width(120.dp), maxLines = 1, overflow = TextOverflow.Ellipsis)
                        if (showStationColumn) Text(item.stationName, fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(95.dp), maxLines = 1, overflow = TextOverflow.Ellipsis)
                        Row(modifier = Modifier.width(95.dp), verticalAlignment = Alignment.CenterVertically) { 
                            Text("№${item.trainNum}", fontWeight = FontWeight.Black, fontFamily = FontFamily.Monospace, fontSize = 12.sp, color = TrainRTheme.NavyDark, maxLines = 1)
                            if (item.isMcd) { 
                                Spacer(modifier = Modifier.width(2.dp)); 
                                Surface(shape = RoundedCornerShape(2.dp), color = Color(0xFFFFE4E6)) { 
                                    // Бейдж МЦД теперь горизонтальный (maxLines = 1)
                                    Text("МЦД", fontSize = 8.sp, fontWeight = FontWeight.Black, color = Color(0xFFBE123C), modifier = Modifier.padding(horizontal = 2.dp, vertical = 1.dp), maxLines = 1) 
                                } 
                            } 
                        }
                        Column(modifier = Modifier.width(135.dp)) { Text(item.routeString, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark, maxLines = 1, overflow = TextOverflow.Ellipsis); Text(item.destinationLabel, fontWeight = FontWeight.SemiBold, fontSize = 9.sp, color = TrainRTheme.RoyalBlue, maxLines = 1) }
                        Text(item.telegramHeader, fontWeight = FontWeight.Bold, fontSize = 10.sp, color = TrainRTheme.TextSub, modifier = Modifier.width(115.dp), maxLines = 1, overflow = TextOverflow.Ellipsis)
                        Text(item.afishaTime, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.ExtraBold, fontSize = 12.sp, color = TrainRTheme.TextSub, modifier = Modifier.width(55.dp))
                        Text(item.telegramTime, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Black, fontSize = 12.sp, color = if (item.isCancelledAtStation) TrainRTheme.RedCancel else TrainRTheme.RoyalBlue, modifier = Modifier.width(55.dp))
                        Text(item.finalChangeText, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = when { item.validationError.isNotEmpty() -> TrainRTheme.RedCancel; item.isCancelledAtStation -> TrainRTheme.RedCancel; item.isAppointedAtStation || item.isAppointedStop -> TrainRTheme.GreenSuccess; else -> TrainRTheme.TextDark }, modifier = Modifier.weight(1f).padding(end = 4.dp), maxLines = 2, overflow = TextOverflow.Ellipsis)
                        Text(item.otmenyFileStatus.ifEmpty { "—" }, fontWeight = FontWeight.Bold, fontSize = 10.sp, color = if (item.hasDiscrepancyWithOtmeny) TrainRTheme.RedCancel else TrainRTheme.TextSub, modifier = Modifier.width(120.dp), maxLines = 1, overflow = TextOverflow.Ellipsis)
                    }
                }
            }
        }
    }
}

@Composable
fun StationAnalyticsReportView(stationName: String, dateFilterLabel: String, toMoscow: List<StationComparisonRow>, fromMoscow: List<StationComparisonRow>, onTrainClick: (StationComparisonRow) -> Unit) {
    var selectedCategory by remember { mutableStateOf(0) }
    val all = (toMoscow + fromMoscow).distinctBy { "${it.trainNum}_${it.direction}_${it.dateGroup}" }
    
    val cancelledTrains = all.filter { it.isCancelledAtStation && !it.isPartialCancel }
    val appointedTrains = all.filter { it.isAppointedAtStation }
    val shifted = all.filter { !it.isCancelledAtStation && !it.isAppointedAtStation && !it.isAppointedStop && (it.isTimeShiftedAtStation || it.finalChangeText.contains("отправится", true)) }
    val discrepancies = all.filter { it.hasDiscrepancyWithOtmeny }

    Column(modifier = Modifier.fillMaxSize().verticalScroll(rememberScrollState()).padding(16.dp)) {
        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
            Column {
                Text("Сводная отчетность по станции: $stationName", fontSize = 18.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                Text("Период выборки: $dateFilterLabel • Нажмите на карточку для детализации списка внизу", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
            }
            if (selectedCategory != 0) {
                FilledTonalButton(onClick = { selectedCategory = 0 }, contentPadding = PaddingValues(horizontal = 10.dp, vertical = 4.dp), modifier = Modifier.height(30.dp)) {
                    Text("Сбросить фильтр (показать все)", fontSize = 11.sp, fontWeight = FontWeight.ExtraBold)
                }
            }
        }

        Spacer(modifier = Modifier.height(10.dp))

        Row(horizontalArrangement = Arrangement.spacedBy(10.dp), modifier = Modifier.fillMaxWidth()) {
            StatSummaryBox("Отменено поездов", cancelledTrains.size, TrainRTheme.RedCancel, selectedCategory == 1, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 1) 0 else 1 }
            StatSummaryBox("Назначено поездов", appointedTrains.size, TrainRTheme.GreenSuccess, selectedCategory == 2, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 2) 0 else 2 }
            StatSummaryBox("Изменено время отпр.", shifted.size, TrainRTheme.GoldAccent, selectedCategory == 3, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 3) 0 else 3 }
            StatSummaryBox("Расхождений с выгрузкой", discrepancies.size, TrainRTheme.NavyDark, selectedCategory == 4, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 4) 0 else 4 }
        }

        Spacer(modifier = Modifier.height(10.dp))

        Surface(modifier = Modifier.fillMaxWidth(), shape = RoundedCornerShape(8.dp), color = Color(0xFFF8FAFC), border = BorderStroke(1.dp, TrainRTheme.BorderCol)) {
            Column(modifier = Modifier.padding(10.dp), verticalArrangement = Arrangement.spacedBy(4.dp)) {
                val cancSummary = if (cancelledTrains.isEmpty()) "$stationName: отменено 0 поездов."
                    else "$stationName: отменено ${cancelledTrains.size} поездов — " + cancelledTrains.joinToString(", ") { "№${it.trainNum} (${it.routeString}) в ${if (it.afishaTime != "-") it.afishaTime else it.telegramTime}" }
                Text(cancSummary, fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.RedCancel)

                val appSummary = if (appointedTrains.isEmpty()) "$stationName: назначено 0 поездов."
                    else "$stationName: назначено ${appointedTrains.size} поездов — " + appointedTrains.joinToString(", ") { "№${it.trainNum} (${it.routeString}) в ${it.telegramTime}" }
                Text(appSummary, fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.GreenSuccess)

                val shiftSummary = if (shifted.isEmpty()) "$stationName: изменено время отправления у 0 поездов."
                    else "$stationName: изменено время у ${shifted.size} поездов (Афиша ➔ Телеграмма) — " + shifted.joinToString(", ") { "№${it.trainNum} (${it.afishaTime} ➔ ${it.telegramTime})" }
                Text(shiftSummary, fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.AmberWarn)
            }
        }

        Spacer(modifier = Modifier.height(12.dp))

        val activeCategoryTitle = when (selectedCategory) {
            1 -> "Детализация: Отмененные поезда (${cancelledTrains.size})"
            2 -> "Детализация: Назначенные поезда (${appointedTrains.size})"
            3 -> "Детализация: Измененные поезда (${shifted.size})"
            4 -> "Детализация: Расхождения с выгрузкой (${discrepancies.size})"
            else -> "Полная детализация по всем изменениям (${all.size})"
        }

        Text(activeCategoryTitle, fontWeight = FontWeight.Black, fontSize = 14.sp, color = TrainRTheme.NavyDark)
        Spacer(modifier = Modifier.height(4.dp))

        Row(modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 6.dp, horizontal = 10.dp), verticalAlignment = Alignment.CenterVertically) {
            Text("ДАТА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(120.dp))
            Text("№ ПОЕЗДА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(95.dp))
            Text("МАРШРУТ СЛЕДОВАНИЯ", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(175.dp))
            Text("АФИША ➔ ТЕЛЕГРАММА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(195.dp))
            Text("ПОЛНЫЙ СТАТУС ИЗМЕНЕНИЙ ПО СТАНЦИИ", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
        }

        Spacer(modifier = Modifier.height(4.dp))

        val displayList = when (selectedCategory) {
            1 -> cancelledTrains
            2 -> appointedTrains
            3 -> shifted
            4 -> discrepancies
            else -> cancelledTrains + appointedTrains + shifted + discrepancies
        }

        if (displayList.isEmpty()) {
            Box(modifier = Modifier.fillMaxWidth().height(100.dp), contentAlignment = Alignment.Center) {
                Text("Нет поездов в выбранной категории отчётности.", fontSize = 12.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub)
            }
        } else {
            Column(modifier = Modifier.fillMaxWidth(), verticalArrangement = Arrangement.spacedBy(5.dp)) {
                displayList.forEach { r ->
                    val accentColor = when {
                        r.isCancelledAtStation -> TrainRTheme.RedCancel
                        r.isAppointedAtStation -> TrainRTheme.GreenSuccess
                        r.hasDiscrepancyWithOtmeny -> TrainRTheme.NavyDark
                        else -> TrainRTheme.GoldAccent
                    }
                    DetailedReportTrainCard(r, accentColor) { onTrainClick(r) }
                }
            }
        }
    }
}

@Composable
fun TrainDossierDialog(row: StationComparisonRow, onDismiss: () -> Unit) {
    val rawTextScrollState = rememberScrollState()
    var selectedStationInDossier by remember { mutableStateOf<String?>(row.stationName) }

    Dialog(onDismissRequest = onDismiss) {
        Card(modifier = Modifier.width(840.dp).heightIn(min = 520.dp, max = 740.dp), shape = RoundedCornerShape(12.dp), colors = CardDefaults.cardColors(containerColor = Color.White), elevation = CardDefaults.cardElevation(8.dp)) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text("ПОЛНЫЙ ОТЧЕТ ПО ПОЕЗДУ №${row.trainNum} ${row.daysFlag}", fontSize = 16.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                            Spacer(modifier = Modifier.width(8.dp))
                            Surface(shape = RoundedCornerShape(4.dp), color = TrainRTheme.RoyalBlue) { Text("СТАНЦИЯ: ${row.stationName.uppercase()}", fontSize = 10.sp, fontWeight = FontWeight.Black, color = Color.White, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp)) }
                        }
                        Text("Маршрут: ${row.routeString} (${row.destinationLabel}) • Основание: ${row.telegramHeader}", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
                    }
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, contentDescription = null) }
                }
                Divider(modifier = Modifier.padding(vertical = 6.dp))
                if (row.validationError.isNotEmpty()) Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFFFF1F2), modifier = Modifier.fillMaxWidth().padding(bottom = 6.dp)) { Text("ОШИБКА: ${row.validationError}", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RedCancel, modifier = Modifier.padding(6.dp)) }
                Text("Период действия: ${row.dateGroup}", fontWeight = FontWeight.ExtraBold, fontSize = 12.sp)
                Text("Время по ст. ${row.stationName}: Афиша ${row.afishaTime} ➔ Телеграмма ${row.telegramTime}", fontWeight = FontWeight.Black, fontSize = 12.sp, color = TrainRTheme.RoyalBlue)
                Text("Итоговый статус: ${row.finalChangeText}", fontWeight = FontWeight.Black, fontSize = 12.sp, color = if (row.isCancelledAtStation) TrainRTheme.RedCancel else TrainRTheme.GreenSuccess)
                if (row.trackSummaryForStation.isNotEmpty()) Text("Изменение пути: ${row.trackSummaryForStation}", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.RoyalBlue)
                if (row.skippedStopsSummary.isNotEmpty()) Text("Отменены остановки: ${row.skippedStopsSummary}", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = TrainRTheme.RedCancel)
                if (row.appointedStopsSummary.isNotEmpty()) Text("Назначены остановки: ${row.appointedStopsSummary}", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = TrainRTheme.GreenSuccess)
                Spacer(modifier = Modifier.height(6.dp))
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Text("Расписание по станциям следования (нажмите на станцию для справки):", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark)
                    Text("Кликните на станцию", fontSize = 10.sp, color = TrainRTheme.RoyalBlue, fontWeight = FontWeight.Bold)
                }
                Spacer(modifier = Modifier.height(3.dp))
                Row(modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(4.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 4.dp, horizontal = 8.dp), verticalAlignment = Alignment.CenterVertically) {
                    Text("СТАНЦИЯ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(180.dp))
                    Text("ПО АФИШЕ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(100.dp))
                    Text("ПО ТЕЛЕГРАММЕ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(140.dp))
                    Text("РАЗНИЦА / СТАТУС СТАНЦИИ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
                }
                val allStKeys = (row.fullAfishaStations.keys + row.fullTelegramStations.keys).distinct()
                LazyColumn(modifier = Modifier.weight(1f).fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF8FAFC)).border(BorderStroke(1.dp, TrainRTheme.BorderCol), RoundedCornerShape(6.dp))) {
                    items(allStKeys) { stKey ->
                        val dispName = displayStationName(stKey)
                        val isCurrentViewSt = (stKey == normalizeStationKey(row.stationName))
                        val isSelectedInModal = (dispName == selectedStationInDossier)
                        val afT = row.fullAfishaStations[stKey] ?: "-"
                        val telPair = row.fullTelegramStations[stKey]
                        val telT = if (telPair != null) { if (telPair.second) telPair.first else "${telPair.first} (без ост.)" } else "-"
                        val diffLabel = if (afT != "-" && telPair != null && telPair.second) calculateMinutesDiffLabel(afT, telPair.first) else ""
                        val statusBadge = when {
                            telPair == null && afT != "-" -> "Отмена по станции"
                            telPair != null && !telPair.second -> "Без остановки"
                            diffLabel.isNotEmpty() -> diffLabel
                            else -> "По графику"
                        }
                        val rowBg = when { isSelectedInModal -> TrainRTheme.GoldAccent.copy(alpha = 0.15f); isCurrentViewSt -> TrainRTheme.RoyalBlue.copy(alpha = 0.08f); else -> Color.Transparent }
                        Row(modifier = Modifier.fillMaxWidth().background(rowBg).clickable { selectedStationInDossier = dispName }.padding(vertical = 4.dp, horizontal = 8.dp), verticalAlignment = Alignment.CenterVertically) {
                            Row(modifier = Modifier.width(180.dp), verticalAlignment = Alignment.CenterVertically) {
                                if (isCurrentViewSt) { Box(modifier = Modifier.size(6.dp).clip(CircleShape).background(TrainRTheme.RoyalBlue)); Spacer(modifier = Modifier.width(4.dp)) }
                                Text(dispName, fontWeight = if (isCurrentViewSt || isSelectedInModal) FontWeight.Black else FontWeight.Bold, fontSize = 11.sp, color = TrainRTheme.NavyDark, maxLines = 1, overflow = TextOverflow.Ellipsis)
                            }
                            Text(afT, fontFamily = FontFamily.Monospace, fontSize = 11.sp, fontWeight = FontWeight.SemiBold, modifier = Modifier.width(100.dp))
                            Text(text = telT, fontFamily = FontFamily.Monospace, fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, color = if (telT != "-" && telT != afT) TrainRTheme.RoyalBlue else TrainRTheme.TextSub, modifier = Modifier.width(140.dp))
                            Surface(shape = RoundedCornerShape(4.dp), color = when { statusBadge.contains("Отмена") || statusBadge.contains("Без ост") -> TrainRTheme.RedCancel.copy(alpha = 0.12f); diffLabel.isNotEmpty() -> TrainRTheme.GoldAccent.copy(alpha = 0.12f); else -> Color(0xFFE2E8F0) }) {
                                Text(text = statusBadge, fontSize = 10.sp, fontWeight = FontWeight.ExtraBold, color = when { statusBadge.contains("Отмена") || statusBadge.contains("Без ост") -> TrainRTheme.RedCancel; diffLabel.isNotEmpty() -> TrainRTheme.GoldAccent; else -> TrainRTheme.TextSub }, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp))
                            }
                        }
                        Divider(color = Color(0xFFE2E8F0))
                    }
                }
                if (selectedStationInDossier != null) {
                    Spacer(modifier = Modifier.height(6.dp))
                    Surface(modifier = Modifier.fillMaxWidth(), shape = RoundedCornerShape(8.dp), color = Color(0xFFEFF6FF), border = BorderStroke(1.dp, Color(0xFFBFDBFE))) {
                        Column(modifier = Modifier.padding(8.dp), verticalArrangement = Arrangement.spacedBy(2.dp)) {
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Icon(Icons.Default.Info, null, tint = TrainRTheme.RoyalBlue, modifier = Modifier.size(15.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("Справка по ст. ${selectedStationInDossier!!.uppercase()}:", fontSize = 11.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                                Spacer(modifier = Modifier.weight(1f))
                                Text(row.telegramHeader, fontSize = 10.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
                            }
                            val normSel = normalizeStationKey(selectedStationInDossier!!)
                            val isSkipped = row.skippedStopsSummary.contains(selectedStationInDossier!!, true)
                            val isAppointed = row.appointedStopsSummary.contains(selectedStationInDossier!!, true)
                            val afTime = row.fullAfishaStations[normSel] ?: "-"
                            val telTime = row.fullTelegramStations[normSel]?.first ?: "-"
                            val noteText = when {
                                isSkipped -> "Станция отменена в телеграмме: поезд проследует без остановки."
                                isAppointed -> "Станции назначена остановка в телеграмме: отправление в $telTime."
                                telTime != "-" && afTime != "-" && telTime != afTime -> "Время отправления изменено: по Афише $afTime ➔ по Телеграмме $telTime ${calculateMinutesDiffLabel(afTime, telTime)}."
                                telTime != "-" -> "Время отправления по телеграмме: $telTime."
                                else -> "Станция следует по расписанию нитки либо отменена на данном участке."
                            }
                            Text(noteText, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.NavyDark)
                            val excerpt = row.rawTelegramText.lines().firstOrNull { it.contains(selectedStationInDossier!!, true) } ?: row.rawTelegramText.take(120)
                            Text("Комментарий из телеграммы: \"$excerpt\"", fontSize = 10.sp, fontStyle = FontStyle.Italic, color = TrainRTheme.TextSub)
                        }
                    }
                }
                Spacer(modifier = Modifier.height(6.dp))
                Text("Полный исходный текст по данному поезду из телеграммы:", fontWeight = FontWeight.ExtraBold, fontSize = 10.sp)
                Spacer(modifier = Modifier.height(2.dp))
                Surface(modifier = Modifier.fillMaxWidth().height(80.dp), shape = RoundedCornerShape(6.dp), color = Color(0xFFF1F5F9), border = BorderStroke(1.dp, TrainRTheme.BorderCol)) {
                    Box(modifier = Modifier.fillMaxSize().padding(6.dp).verticalScroll(rawTextScrollState)) {
                        Text(row.rawTelegramText, fontSize = 11.sp, lineHeight = 15.sp, fontWeight = FontWeight.SemiBold, color = TrainRTheme.NavyDark)
                    }
                }
            }
        }
    }
}

@Composable
fun YandexDiscrepancyView(rows: List<YandexDiscrepancyRow>, onRowClick: (YandexDiscrepancyRow) -> Unit) {
    Column(modifier = Modifier.fillMaxSize().padding(8.dp)) {
        Text("Сверка с Яндекс.Расписанием:", fontWeight = FontWeight.Black, fontSize = 14.sp, color = TrainRTheme.NavyDark)
        Spacer(modifier = Modifier.height(6.dp))
        Row(modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 6.dp, horizontal = 8.dp), verticalAlignment = Alignment.CenterVertically) {
            Text("№", fontWeight = FontWeight.Black, fontSize = 11.sp, modifier = Modifier.width(80.dp))
            Text("МАРШРУТ", fontWeight = FontWeight.Black, fontSize = 11.sp, modifier = Modifier.width(200.dp))
            Text("ЯНДЕКС", fontWeight = FontWeight.Black, fontSize = 11.sp, modifier = Modifier.width(65.dp))
            Text("ОЖИДАЕТСЯ", fontWeight = FontWeight.Black, fontSize = 11.sp, modifier = Modifier.width(80.dp))
            Text("ОПИСАНИЕ РАСХОЖДЕНИЯ", fontWeight = FontWeight.Black, fontSize = 11.sp, modifier = Modifier.weight(1f))
        }
        Spacer(modifier = Modifier.height(2.dp))
        if (rows.isEmpty()) {
            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) { Text("Загрузите Excel Яндекса — расхождений нет либо файл не загружен.", fontSize = 12.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub) }
        } else {
            LazyColumn(modifier = Modifier.fillMaxSize()) {
                items(rows) { r ->
                    val bg = when (r.severity) { YandexDiscrepancySeverity.CRITICAL_CANCEL_MISSED -> Color(0xFFFECDD3); YandexDiscrepancySeverity.TIME_MISMATCH -> Color(0xFFFEF3C7); YandexDiscrepancySeverity.STOP_MISMATCH -> Color(0xFFFED7AA); YandexDiscrepancySeverity.INFO -> Color(0xFFE0F2FE) }
                    Row(modifier = Modifier.fillMaxWidth().padding(vertical = 1.dp).clip(RoundedCornerShape(4.dp)).background(bg).border(1.dp, TrainRTheme.BorderCol, RoundedCornerShape(4.dp)).clickable { onRowClick(r) }.padding(vertical = 4.dp, horizontal = 8.dp), verticalAlignment = Alignment.CenterVertically) {
                        Text("№${r.trainNum}", fontWeight = FontWeight.Black, fontFamily = FontFamily.Monospace, fontSize = 12.sp, modifier = Modifier.width(80.dp))
                        Text(r.route, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.width(200.dp), maxLines = 1, overflow = TextOverflow.Ellipsis)
                        Text(r.yandexTime, fontFamily = FontFamily.Monospace, fontSize = 11.sp, fontWeight = FontWeight.Black, modifier = Modifier.width(65.dp))
                        Text(r.actualExpectedTime, fontFamily = FontFamily.Monospace, fontSize = 11.sp, fontWeight = FontWeight.Black, color = TrainRTheme.RoyalBlue, modifier = Modifier.width(80.dp))
                        Text(r.discrepancyDescription, fontSize = 10.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.RedCancel, modifier = Modifier.weight(1f), maxLines = 2, overflow = TextOverflow.Ellipsis)
                    }
                }
            }
        }
    }
}

@Composable
fun YandexDiscrepancyDetailDialog(row: YandexDiscrepancyRow, onDismiss: () -> Unit) {
    Dialog(onDismissRequest = onDismiss) {
        Card(modifier = Modifier.width(720.dp).heightIn(max = 560.dp), shape = RoundedCornerShape(12.dp)) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Text("Расхождение с Яндексом: №${row.trainNum}", fontSize = 16.sp, fontWeight = FontWeight.Black)
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, null) }
                }
                Divider(modifier = Modifier.padding(vertical = 6.dp))
                Text("Маршрут: ${row.route}", fontSize = 12.sp, fontWeight = FontWeight.Bold)
                Text("Время в Яндексе: ${row.yandexTime}", fontSize = 12.sp, fontFamily = FontFamily.Monospace)
                Text("Ожидаемое время: ${row.actualExpectedTime}", fontSize = 12.sp, fontFamily = FontFamily.Monospace, color = TrainRTheme.RoyalBlue)
                Spacer(modifier = Modifier.height(6.dp))
                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFFEF3C7), modifier = Modifier.fillMaxWidth()) {
                    Text(row.discrepancyDescription, fontSize = 12.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.AmberWarn, modifier = Modifier.padding(8.dp))
                }
            }
        }
    }
}

// ==================== 9. ГЛАВНЫЙ ЭКРАН ====================

@OptIn(androidx.compose.foundation.layout.ExperimentalLayoutApi::class)
@Composable
fun TrainRApp() {
    val cachedAfisha = remember { getCachedAfishaFile() }
    var afishaFile by remember { mutableStateOf<File?>(if (cachedAfisha.exists()) cachedAfisha else null) }
    var telegramFiles by remember { mutableStateOf<List<File>>(emptyList()) }; var otmenyFiles by remember { mutableStateOf<List<File>>(emptyList()) }; var manualFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var yandexFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var afishaDb by remember { mutableStateOf<Map<String, AfishaTrainInfo>>(emptyMap()) }; var telegramEvents by remember { mutableStateOf<List<TelegramTrainEvent>>(emptyList()) }; var otmenyMap by remember { mutableStateOf<Map<String, String>>(emptyMap()) }; var manualMap by remember { mutableStateOf<Map<String, String>>(emptyMap()) }
    var yandexTrains by remember { mutableStateOf<List<YandexParsedTrain>>(emptyList()) }
    var yandexDiscrepancies by remember { mutableStateOf<List<YandexDiscrepancyRow>>(emptyList()) }
    var inspectedYandex by remember { mutableStateOf<YandexDiscrepancyRow?>(null) }
    var globalErrors by remember { mutableStateOf<List<GlobalError>>(emptyList()) }
    var ignoredErrors by remember { mutableStateOf<Set<String>>(emptySet()) }
    var availableDates by remember { mutableStateOf<List<SingleDateFilterOption>>(emptyList()) }; var selectedDates by remember { mutableStateOf<Set<SingleDateFilterOption>>(emptySet()) }
    var currentStation by remember { mutableStateOf("") }; var selectedStationsForExport by remember { mutableStateOf<Set<String>>(emptySet()) }
    var toMoscowList by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }; var fromMoscowList by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }
    var activeTab by remember { mutableStateOf(0) }; var searchQuery by remember { mutableStateOf("") }; var globalSearchResults by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }; var inspectedTrainRow by remember { mutableStateOf<StationComparisonRow?>(null) }
    var isBusy by remember { mutableStateOf(false) }; var statusText by remember { mutableStateOf("Загрузите Афишу (.xlsx), Телеграммы (.docx/.pdf) или файл «Отмены».") }
    var activeTrainFilter by remember { mutableStateOf(TrainFilter.ALL) }
    var showCalendarDialog by remember { mutableStateOf(false) }
    val scope = rememberCoroutineScope(); var recalcJob by remember { mutableStateOf<Job?>(null) }

    LaunchedEffect(afishaFile) { if (afishaFile != null && afishaFile!!.exists()) withContext(Dispatchers.IO) { val loaded = loadAfishaDatabase(afishaFile); afishaDb = loaded; statusText = "Базовая Афиша активна (распознано ${loaded.size} ниток поездов)." } }
    LaunchedEffect(availableDates) { if (selectedDates.isNotEmpty() && !selectedDates.all { it in availableDates }) selectedDates = selectedDates.filter { it in availableDates }.toSet() }

    fun openFileDialog(title: String, multi: Boolean, filterExt: String? = null): List<File> {
        val dlg = FileDialog(null as Frame?, title, FileDialog.LOAD).apply { isMultipleMode = multi; if (filterExt != null) setFilenameFilter { _, name -> name.endsWith(filterExt, true) }; isVisible = true }
        val res = dlg.files.toList(); dlg.dispose(); return res
    }

    fun recalcYandex() {
        yandexTrains = yandexFiles.flatMap { parseYandexExcel(it) }
        yandexDiscrepancies = buildYandexDiscrepancies(yandexTrains, telegramEvents, afishaDb)
    }

    fun recalculateStation(stName: String, dates: Set<SingleDateFilterOption> = selectedDates) {
        recalcJob?.cancel()
        recalcJob = scope.launch {
            isBusy = true; delay(AppConstants.REPARSE_DEBOUNCE_MS)
            statusText = if (stName == "ВСЕ") "Сверка ВСЕХ станций..." else "Сверка расписания по ст. $stName..."
            val res = withContext(Dispatchers.IO) {
                val telsPair = TelegramParseCache.get(telegramFiles); val otm = OtmenyParseCache.get(otmenyFiles)
                val datesList = extractAvailableDateOptions(telsPair.first)
                if (stName == "ВСЕ") {
                    val allTo = mutableListOf<StationComparisonRow>(); val allFrom = mutableListOf<StationComparisonRow>()
                    POLYGON_STATIONS.forEach { st ->
                        val (t, f) = buildStationComparison(st, afishaDb, telsPair.first, otm, manualMap, dates, ignoredErrors)
                        allTo.addAll(t.map { it.copy(stationName = st) }); allFrom.addAll(f.map { it.copy(stationName = st) })
                    }
                    StationCalcResult(allTo.distinctBy{it.trainNum + it.finalChangeText}, allFrom.distinctBy{it.trainNum + it.finalChangeText}, datesList, telsPair.first, otm, telsPair.second)
                } else {
                    val (t, f) = buildStationComparison(stName, afishaDb, telsPair.first, otm, manualMap, dates, ignoredErrors)
                    StationCalcResult(t, f, datesList, telsPair.first, otm, telsPair.second)
                }
            }
            telegramEvents = res.telegramEvents; otmenyMap = res.otmenyMap; availableDates = res.availableDates; toMoscowList = res.toMoscow; fromMoscowList = res.fromMoscow; globalErrors = res.errors
            recalcYandex()
            isBusy = false; statusText = if (stName == "ВСЕ") "Сверка всех станций завершена. Найдено изменений: ${res.toMoscow.size + res.fromMoscow.size}." else "Станция $stName: До Москвы — ${res.toMoscow.size}, От Москвы — ${res.fromMoscow.size}. Ошибок: ${res.errors.size}."
        }
    }

    fun runGlobalTrainSearch(query: String) {
        val q = query.trim(); if (q.isEmpty()) return
        scope.launch {
            isBusy = true; statusText = "Глобальный поиск поезда №$q..."
            val found = withContext(Dispatchers.IO) {
                val tels = if (telegramEvents.isEmpty() && telegramFiles.isNotEmpty()) TelegramParseCache.get(telegramFiles).first else telegramEvents
                val otm = if (otmenyMap.isEmpty() && otmenyFiles.isNotEmpty()) OtmenyParseCache.get(otmenyFiles) else otmenyMap
                val results = mutableListOf<StationComparisonRow>()
                
                fun numMatches(trainNum: String): Boolean = trainNum.contains(q, true) || trainNumLookupKeys(trainNum).any { it.contains(q, true) }

                for (st in POLYGON_STATIONS) { 
                    val (t, f) = buildStationComparison(st, afishaDb, tels, otm, manualMap, selectedDates, ignoredErrors)
                    results.addAll((t + f).filter { numMatches(it.trainNum) }) 
                }
                
                if (results.isEmpty()) {
                    val matchingEv = tels.firstOrNull { numMatches(it.trainNum) }
                    afishaDb.values.filter { numMatches(it.trainNum) }.forEach { af ->
                        results.add(StationComparisonRow(
                            matchingEv?.calendarDays?.fallbackLabel ?: "Афиша", 1, emptyList(), "", af.trainNum, af.daysFlag, 
                            af.routeString, af.destinationLabel, af.isMcd, af.direction, "Все станции", "Без изменений", 
                            af.stationTimes.values.firstOrNull() ?: "-", "-", "", "Следует расписанием Афиши", 
                            false, false, false, false, false, false, "", "", "", false, af.stationTimes, emptyMap(), "", ""
                        ))
                    }
                }
                results.distinctBy { "${it.trainNum}_${it.stationName}_${it.direction}" }
            }
            globalSearchResults = found; activeTab = 3; isBusy = false; statusText = "Поиск завершен: найдено — ${found.size}."
        }
    }

    fun filterByTab(list: List<StationComparisonRow>): List<StationComparisonRow> {
        var filtered = if (searchQuery.isBlank()) list else list.filter { it.trainNum.contains(searchQuery.trim(), true) || trainNumLookupKeys(it.trainNum).any { k -> k.contains(searchQuery.trim(), true) } }
        filtered = when (activeTrainFilter) {
            TrainFilter.MCD -> filtered.filter { it.isMcd }
            TrainFilter.EXPRESS -> filtered.filter { isTrainExpress(it.trainNum) }
            TrainFilter.ALL -> filtered
        }
        return filtered
    }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = TrainRTheme.BgMain) {
            Column(modifier = Modifier.fillMaxSize().padding(8.dp)) {
                Card(modifier = Modifier.fillMaxWidth(), shape = RoundedCornerShape(10.dp), colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg), elevation = CardDefaults.cardElevation(2.dp)) {
                    Row(modifier = Modifier.fillMaxWidth().padding(horizontal = 10.dp, vertical = 6.dp), verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        VintageTrainLogo(Modifier.size(36.dp))
                        Column(modifier = Modifier.width(180.dp)) {
                            Row(verticalAlignment = Alignment.CenterVertically) { Text("Train-R", fontSize = 18.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark); Spacer(modifier = Modifier.width(6.dp)); Surface(shape = RoundedCornerShape(4.dp), color = Color(0xFFFEF3C7)) { Text("МТППК", fontSize = 8.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.GoldAccent, modifier = Modifier.padding(horizontal = 4.dp, vertical = 1.dp)) } }
                            Text(statusText, fontSize = 9.sp, fontWeight = FontWeight.SemiBold, color = TrainRTheme.TextSub, maxLines = 1, overflow = TextOverflow.Ellipsis)
                        }
                        CompactFileBadge("1. АФИША (.XLSX)", if (afishaFile != null) "Ниток: ${afishaDb.size}" else "Не загружена", afishaFile != null, TrainRTheme.GreenSuccess, Modifier.weight(1f), { val f = openFileDialog("Выберите файл Афиши (.xlsx)", false); if (f.isNotEmpty()) { try { FileInputStream(f.first()).use { inp -> FileOutputStream(cachedAfisha).use { out -> inp.copyTo(out) } }; afishaFile = cachedAfisha; afishaDb = loadAfishaDatabase(cachedAfisha) } catch (e: Exception) { afishaFile = f.first(); afishaDb = loadAfishaDatabase(f.first()) }; if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates) } }, { if (cachedAfisha.exists()) cachedAfisha.delete(); afishaFile = null; afishaDb = emptyMap(); statusText = "Афиша удалена." })
                        CompactFileBadge("2. ТЕЛЕГРАММЫ (${telegramFiles.size})", if (telegramFiles.isNotEmpty()) telegramFiles.joinToString { it.nameWithoutExtension.take(8) } else "Word / PDF", telegramFiles.isNotEmpty(), TrainRTheme.RoyalBlue, Modifier.weight(1f), { val f = openFileDialog("Выберите одну или несколько телеграмм", true); if (f.isNotEmpty()) { telegramFiles = telegramFiles + f; if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates) } }, { telegramFiles = emptyList(); telegramEvents = emptyList(); availableDates = emptyList(); selectedDates = emptySet(); statusText = "Телеграммы очищены." })
                        CompactFileBadge("3. ФАЙЛ «ОТМЕНЫ» (${otmenyFiles.size})", if (otmenyFiles.isNotEmpty()) "Загружено: ${otmenyFiles.size}" else "TXT / DOCX", otmenyFiles.isNotEmpty(), Color(0xFF7C3AED), Modifier.weight(1f), { val f = openFileDialog("Выберите файлы выгрузки «Отмены» (.txt, .docx)", true); if (f.isNotEmpty()) { otmenyFiles = f; if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates) } }, { otmenyFiles = emptyList(); otmenyMap = emptyMap(); statusText = "Отмены очищены." })
                        CompactFileBadge("4. СВЕРКА (${manualFiles.size})", if (manualFiles.isNotEmpty()) "Загружено: ${manualFiles.size}" else "Своя выгрузка", manualFiles.isNotEmpty(), Color(0xFF0EA5E9), Modifier.weight(1f), { val f = openFileDialog("Выберите файл(ы) для сверки (.txt, .docx)", true); if (f.isNotEmpty()) { manualFiles = f; scope.launch { isBusy = true; manualMap = withContext(Dispatchers.IO) { parseManualExport(f) }; if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates) } } }, { manualFiles = emptyList(); manualMap = emptyMap(); statusText = "Сверка выгрузок очищена." })
                        CompactFileBadge("5. ЯНДЕКС (.XLSX)", if (yandexFiles.isNotEmpty()) "Строк: ${yandexTrains.size}" else "Excel", yandexFiles.isNotEmpty(), TrainRTheme.YandexRed, Modifier.weight(1f), {
                            val f = openFileDialog("Выберите Excel Яндекса", false, ".xlsx")
                            if (f.isNotEmpty()) {
                                yandexFiles = f
                                scope.launch {
                                    isBusy = true
                                    recalcYandex()
                                    isBusy = false
                                    activeTab = 5
                                }
                            }
                        }, { yandexFiles = emptyList(); yandexTrains = emptyList(); yandexDiscrepancies = emptyList() })
                    }
                }
                Spacer(modifier = Modifier.height(4.dp))
                Card(modifier = Modifier.fillMaxWidth(), shape = RoundedCornerShape(10.dp), colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg), elevation = CardDefaults.cardElevation(1.dp)) {
                    Column(modifier = Modifier.fillMaxWidth().padding(8.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                            Text("Станция:", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark)
                            FilledTonalButton(onClick = { selectedStationsForExport = if (selectedStationsForExport.size == POLYGON_STATIONS.size) emptySet() else POLYGON_STATIONS.toSet() }, contentPadding = PaddingValues(horizontal = 6.dp, vertical = 2.dp), modifier = Modifier.height(26.dp), shape = RoundedCornerShape(6.dp)) { Text(if (selectedStationsForExport.size == POLYGON_STATIONS.size) "Снять галочки" else "Выбрать все", fontSize = 10.sp, fontWeight = FontWeight.Bold) }
                            Spacer(modifier = Modifier.weight(1f))
                            Button(onClick = { recalculateStation("ВСЕ") }, enabled = !isBusy, colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF0F766E)), contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp), modifier = Modifier.height(30.dp), shape = RoundedCornerShape(6.dp)) { Icon(Icons.Default.Refresh, null, modifier = Modifier.size(14.dp)); Spacer(modifier = Modifier.width(4.dp)); Text("Сверить ВСЁ", fontSize = 11.sp, fontWeight = FontWeight.Bold, maxLines = 1) }
                            Button(onClick = { if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates) }, enabled = !isBusy && currentStation.isNotEmpty(), colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RoyalBlue), contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp), modifier = Modifier.height(30.dp), shape = RoundedCornerShape(6.dp)) { Icon(Icons.Default.Refresh, null, modifier = Modifier.size(14.dp)); Spacer(modifier = Modifier.width(4.dp)); Text("Сверить: ${if (currentStation.isEmpty()) "Выберите ст." else currentStation}", fontSize = 11.sp, fontWeight = FontWeight.Bold, maxLines = 1, overflow = TextOverflow.Ellipsis) }
                            Button(onClick = { scope.launch { isBusy = true; statusText = "Генерация выгрузок..."; withContext(Dispatchers.IO) { val pTels = if (telegramEvents.isEmpty()) TelegramParseCache.get(telegramFiles).first else telegramEvents; val pOtm = if (otmenyMap.isEmpty()) OtmenyParseCache.get(otmenyFiles) else otmenyMap; selectedStationsForExport.forEach { st -> val (tM, fM) = buildStationComparison(st, afishaDb, pTels, pOtm, manualMap, selectedDates, ignoredErrors); exportStationFilesToFolders(st, tM, fM, selectedDates.joinToString(", ") { it.displayLabel }) } }; isBusy = false; statusText = "ГОТОВО! Выгрузки сохранены." } }, enabled = !isBusy && selectedStationsForExport.isNotEmpty(), colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RedCancel), contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp), modifier = Modifier.height(30.dp), shape = RoundedCornerShape(6.dp)) { Icon(Icons.Default.Share, null, modifier = Modifier.size(14.dp)); Spacer(modifier = Modifier.width(4.dp)); Text("Генерировать (${selectedStationsForExport.size})", fontSize = 11.sp, fontWeight = FontWeight.Bold, maxLines = 1) }
                        }
                        Spacer(modifier = Modifier.height(6.dp))
                        androidx.compose.foundation.layout.FlowRow(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp), verticalArrangement = Arrangement.spacedBy(3.dp)) {
                            POLYGON_STATIONS.forEach { st -> StationChipItem(st, st in selectedStationsForExport, currentStation == st, Modifier.width(120.dp), { selectedStationsForExport = if (st in selectedStationsForExport) selectedStationsForExport - st else selectedStationsForExport + st }, { currentStation = if (currentStation == st) "" else st; recalculateStation(if (currentStation == st) "" else st, selectedDates) }) }
                        }
                        if (availableDates.isNotEmpty()) {
                            Spacer(modifier = Modifier.height(6.dp)); Divider(color = Color(0xFFE2E8F0)); Spacer(modifier = Modifier.height(6.dp))
                            Row(verticalAlignment = Alignment.CenterVertically) { Icon(Icons.Default.DateRange, null, tint = TrainRTheme.RoyalBlue, modifier = Modifier.size(14.dp)); Spacer(modifier = Modifier.width(4.dp)); Text("Фильтр по дате:", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark)
                                Button(onClick = { showCalendarDialog = true }, contentPadding = PaddingValues(horizontal = 6.dp, vertical = 2.dp), modifier = Modifier.height(26.dp).padding(start = 8.dp), shape = RoundedCornerShape(6.dp)) { Icon(Icons.Default.DateRange, null, modifier = Modifier.size(12.dp)); Spacer(modifier = Modifier.width(2.dp)); Text(if (selectedDates.isEmpty()) "Все дни" else "Выбрано: ${selectedDates.size}", fontSize = 11.sp, fontWeight = FontWeight.Bold) }
                                if (selectedDates.isNotEmpty()) { Spacer(modifier = Modifier.width(6.dp)); Text(selectedDates.joinToString(", ") { it.displayLabel }, fontSize = 11.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.SemiBold, maxLines = 1, overflow = TextOverflow.Ellipsis) }
                            }
                        }
                    }
                }
                Spacer(modifier = Modifier.height(4.dp))
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    TabRow(selectedTabIndex = if (activeTab > 5) 0 else activeTab, containerColor = TrainRTheme.CardBg, contentColor = TrainRTheme.RoyalBlue, modifier = Modifier.width(900.dp).height(38.dp).clip(RoundedCornerShape(8.dp))) {
                        Tab(selected = activeTab == 0, onClick = { activeTab = 0 }, text = { Text("НА МОСКВУ / ИППОДРОМ (${toMoscowList.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, maxLines = 1) })
                        Tab(selected = activeTab == 1, onClick = { activeTab = 1 }, text = { Text("ОТ МОСКВЫ / ИППОДРОМА (${fromMoscowList.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, maxLines = 1) })
                        Tab(selected = activeTab == 2, onClick = { activeTab = 2 }, text = { Text("ОТЧЕТ", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.GreenSuccess, maxLines = 1) })
                        Tab(selected = activeTab == 3, onClick = { activeTab = 3 }, text = { Text("ПОИСК (${globalSearchResults.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.GoldAccent, maxLines = 1) })
                        Tab(selected = activeTab == 4, onClick = { activeTab = 4 }, text = { Text("ОШИБКИ (${globalErrors.count { !it.isIgnored }})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.RedCancel, maxLines = 1) })
                        Tab(selected = activeTab == 5, onClick = { activeTab = 5 }, text = { Text("ЯНДЕКС (${yandexDiscrepancies.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.YandexRed, maxLines = 1) })
                    }
                    Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                        OutlinedTextField(value = searchQuery, onValueChange = { searchQuery = it }, placeholder = { Text("№ поезда...", fontSize = 12.sp) }, modifier = Modifier.width(180.dp).height(42.dp), singleLine = true, textStyle = LocalTextStyle.current.copy(fontSize = 13.sp, fontWeight = FontWeight.Bold), shape = RoundedCornerShape(8.dp), trailingIcon = { if (searchQuery.isNotEmpty()) IconButton(onClick = { searchQuery = "" }) { Icon(Icons.Default.Clear, "Сброс", modifier = Modifier.size(16.dp)) } })
                        Button(onClick = { runGlobalTrainSearch(searchQuery) }, colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark), modifier = Modifier.height(38.dp), shape = RoundedCornerShape(8.dp)) { Icon(Icons.Default.Search, null, modifier = Modifier.size(16.dp)); Spacer(modifier = Modifier.width(4.dp)); Text("Найти", fontSize = 12.sp, fontWeight = FontWeight.Bold) }
                    }
                }
                Spacer(modifier = Modifier.height(4.dp))
                Card(modifier = Modifier.weight(1f), shape = RoundedCornerShape(10.dp), colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg), elevation = CardDefaults.cardElevation(2.dp)) {
                    when (activeTab) {
                        0, 1 -> {
                            val baseList = if (activeTab == 0) toMoscowList else fromMoscowList
                            Column(modifier = Modifier.fillMaxSize()) {
                                Row(modifier = Modifier.fillMaxWidth().padding(horizontal = 8.dp, vertical = 4.dp), horizontalArrangement = Arrangement.spacedBy(4.dp), verticalAlignment = Alignment.CenterVertically) {
                                    Text("Фильтр:", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub)
                                    FilterChipItem("Все", activeTrainFilter == TrainFilter.ALL) { activeTrainFilter = TrainFilter.ALL }
                                    FilterChipItem("МЦД", activeTrainFilter == TrainFilter.MCD) { activeTrainFilter = TrainFilter.MCD }
                                    FilterChipItem("Экспрессы", activeTrainFilter == TrainFilter.EXPRESS) { activeTrainFilter = TrainFilter.EXPRESS }
                                }
                                TrainTableSection(rows = filterByTab(baseList), emptyMessage = "Нет изменений по ст. $currentStation.", onRowClick = { inspectedTrainRow = it })
                            }
                        }
                        2 -> StationAnalyticsReportView(currentStation, selectedDates.joinToString(", ") { it.displayLabel }.ifEmpty { "Все даты" }, toMoscowList, fromMoscowList, { inspectedTrainRow = it })
                        3 -> TrainTableSection(rows = filterByTab(globalSearchResults), emptyMessage = "Введите номер поезда и нажмите «Найти».", showStationColumn = true, onRowClick = { inspectedTrainRow = it })
                        4 -> GlobalErrorsView(globalErrors, onIgnore = { errId -> globalErrors = globalErrors.map { if (it.id == errId) it.copy(isIgnored = true) else it }; if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates) })
                        5 -> YandexDiscrepancyView(yandexDiscrepancies, { inspectedYandex = it })
                    }
                }
                Spacer(modifier = Modifier.height(2.dp))
                if (isBusy) {
                    LinearProgressIndicator(modifier = Modifier.fillMaxWidth().height(3.dp).clip(RoundedCornerShape(2.dp)), color = TrainRTheme.RoyalBlue, trackColor = Color(0xFFE2E8F0))
                } else {
                    Spacer(modifier = Modifier.height(3.dp))
                }
            }
            if (showCalendarDialog) {
                CalendarDatePickerDialog(
                    availableDates = availableDates,
                    initialSelected = selectedDates,
                    onDismiss = { showCalendarDialog = false },
                    onConfirm = { newDates ->
                        selectedDates = newDates
                        showCalendarDialog = false
                        if (currentStation.isNotEmpty()) recalculateStation(currentStation, newDates)
                    }
                )
            }
            if (inspectedTrainRow != null) TrainDossierDialog(row = inspectedTrainRow!!, onDismiss = { inspectedTrainRow = null })
            if (inspectedYandex != null) YandexDiscrepancyDetailDialog(row = inspectedYandex!!, onDismiss = { inspectedYandex = null })
        }
    }
}

@Composable
fun GlobalErrorsView(errors: List<GlobalError>, onIgnore: (String) -> Unit) {
    Column(modifier = Modifier.fillMaxSize().padding(12.dp)) {
        Text("Валидация и ошибки анализа", fontSize = 16.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
        Text("Здесь отображаются конфликты в телеграммах. Если ошибка ложная, вы можете её проигнорировать.", fontSize = 10.sp, color = TrainRTheme.TextSub)
        Spacer(modifier = Modifier.height(8.dp))
        val activeErrors = errors.filter { !it.isIgnored }
        if (activeErrors.isEmpty()) {
            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                Text("Ошибок не обнаружено. Все расписания корректны.", fontSize = 14.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.GreenSuccess)
            }
        } else {
            LazyColumn(verticalArrangement = Arrangement.spacedBy(6.dp)) {
                items(activeErrors) { err ->
                    Surface(shape = RoundedCornerShape(8.dp), color = Color(0xFFFFF1F2), border = BorderStroke(1.dp, Color(0xFFFECDD3)), modifier = Modifier.fillMaxWidth()) {
                        Row(modifier = Modifier.padding(8.dp), verticalAlignment = Alignment.CenterVertically) {
                            Icon(Icons.Default.Warning, null, tint = TrainRTheme.RedCancel, modifier = Modifier.size(20.dp))
                            Spacer(modifier = Modifier.width(8.dp))
                            Column(modifier = Modifier.weight(1f)) {
                                Text("Поезд №${err.trainNum} • ${err.header}", fontSize = 13.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                                Text(err.reason, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RedCancel)
                            }
                            Spacer(modifier = Modifier.width(8.dp))
                            OutlinedButton(onClick = { onIgnore(err.id) }, contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp), modifier = Modifier.height(28.dp), shape = RoundedCornerShape(6.dp), border = BorderStroke(1.dp, TrainRTheme.BorderCol)) {
                                Text("Игнорировать", fontSize = 10.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub)
                            }
                        }
                    }
                }
            }
        }
    }
}

@OptIn(androidx.compose.foundation.layout.ExperimentalLayoutApi::class)
@Composable
fun CalendarDatePickerDialog(
    availableDates: List<SingleDateFilterOption>,
    initialSelected: Set<SingleDateFilterOption>,
    onDismiss: () -> Unit,
    onConfirm: (Set<SingleDateFilterOption>) -> Unit
) {
    val currentSelection = remember { mutableStateOf(initialSelected) }
    
    val months = availableDates.groupBy { it.monthNum to it.monthName }.toList().sortedBy { it.first.first }

    Dialog(onDismissRequest = onDismiss) {
        Card(modifier = Modifier.width(700.dp).heightIn(max = 700.dp), shape = RoundedCornerShape(12.dp)) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                Text("Выбор дат", fontSize = 18.sp, fontWeight = FontWeight.Black)
                Spacer(modifier = Modifier.height(8.dp))
                
                LazyColumn(modifier = Modifier.weight(1f)) {
                    items(months) { (monthPair, datesInMonth) ->
                        val (mNum, mName) = monthPair
                        Text("Месяц: $mName", fontSize = 14.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue, modifier = Modifier.padding(vertical = 8.dp))
                        
                        androidx.compose.foundation.layout.FlowRow(
                            modifier = Modifier.fillMaxWidth(), 
                            maxItemsInEachRow = 7,
                            horizontalArrangement = Arrangement.spacedBy(4.dp), 
                            verticalArrangement = Arrangement.spacedBy(4.dp)
                        ) {
                            for (day in 1..31) {
                                val dateOpt = availableDates.find { it.day == day && it.monthNum == mNum }
                                val isAvailable = dateOpt != null
                                val isSel = dateOpt != null && dateOpt in currentSelection.value
                                
                                Surface(
                                    modifier = Modifier.size(44.dp).clickable(enabled = isAvailable) {
                                        if (isAvailable && dateOpt != null) {
                                            currentSelection.value = if (isSel) currentSelection.value - dateOpt else currentSelection.value + dateOpt
                                        }
                                    },
                                    shape = RoundedCornerShape(8.dp),
                                    color = when {
                                        isSel -> TrainRTheme.GoldAccent
                                        isAvailable -> Color(0xFFE2E8F0)
                                        else -> Color.Transparent
                                    },
                                    border = if (isAvailable && !isSel) BorderStroke(1.dp, TrainRTheme.BorderCol) else null
                                ) {
                                    Box(contentAlignment = Alignment.Center) {
                                        Text(
                                            text = day.toString(), 
                                            color = if (isSel) Color.White else if (isAvailable) TrainRTheme.NavyDark else Color.LightGray,
                                            fontWeight = if (isAvailable) FontWeight.Bold else FontWeight.Normal
                                        )
                                    }
                                }
                            }
                        }
                    }
                }
                
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween) {
                    TextButton(onClick = { currentSelection.value = emptySet() }) { Text("Очистить все", color = TrainRTheme.RedCancel) }
                    Button(onClick = { onConfirm(currentSelection.value) }) { Text("Применить") }
                }
            }
        }
    }
}

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "Train-R — Интеллектуальный ревизор пригородных расписаний",
        state = WindowState(width = 1360.dp, height = 860.dp)
    ) {
        TrainRApp()
    }
}
