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
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.awt.FileDialog
import java.awt.Frame
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.nio.charset.Charset
import java.text.SimpleDateFormat
import java.time.YearMonth
import java.util.Calendar

// ==================== 1. КОНСТАНТЫ И МОДЕЛИ ====================

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
enum class YandexDiscrepancySeverity { CRITICAL_CANCEL_MISSED, TIME_MISMATCH, STOP_MISMATCH, INFO }
enum class MainNavTab { VYGRUZKI, REPORTS_SVERKA, SCHEDULE_YANDEX }

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

// ==================== 2. СТИЛИ И ПАЛИТРА ИНТЕРФЕЙСА (РИС. 2) ====================

object TrainRTheme {
    val SidebarNavy = Color(0xFF0F172A)
    // Compatibility alias used by table headers and row labels.
    val NavyDark = SidebarNavy
    val SidebarActive = Color(0xFF2563EB)
    val BgMain = Color(0xFFF8FAFC)
    val CardBg = Color(0xFFFFFFFF)
    val RoyalBlue = Color(0xFF1D4ED8)
    val ActionBlue = Color(0xFF2563EB)
    val McdCoral = Color(0xFFBE123C)
    val GoldAccent = Color(0xFFD97706)
    val GreenSuccess = Color(0xFF059669)
    val RedCancel = Color(0xFFDC2626)
    val AmberWarn = Color(0xFFB45309)
    val BorderCol = Color(0xFFE2E8F0)
    val TextDark = Color(0xFF0F172A)
    val TextSub = Color(0xFF64748B)
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

// ==================== 3. БАЗОВЫЕ УТИЛИТЫ И ЧТЕНИЕ ФАЙЛОВ ====================

fun getCachedAfishaFile(): File {
    val dir = File(System.getProperty("user.home"), ".train_r_system")
    if (!dir.exists()) dir.mkdirs()
    return File(dir, "cached_afisha.xlsx")
}

private fun logWarn(msg: String, e: Throwable? = null) {
    System.err.println("[TrainR/WARN] $msg")
    e?.printStackTrace()
}

// Надежное чтение текста с автоопределением кодировок (UTF-8, Windows-1251)
fun readDocumentTextAuto(file: File): String {
    if (!file.exists()) return ""
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
            else -> {
                val bytes = file.readBytes()
                try {
                    val s = String(bytes, Charsets.UTF_8)
                    if (s.contains("") && !s.contains("Крюково") && !s.contains("Москва")) {
                        String(bytes, Charset.forName("windows-1251"))
                    } else s
                } catch (_: Exception) {
                    String(bytes, Charset.forName("windows-1251"))
                }
            }
        }
    } catch (e: Exception) {
        logWarn("Ошибка чтения ${file.name}", e)
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

// ==================== 4. ПАРСЕР АФИШИ (EXCEL) ====================

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

// ==================== 5. ПАРСЕР ТЕЛЕГРАММ ====================

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
    val rawText = readDocumentTextAuto(file)
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

// ==================== 6. ПАРСЕР СВОДНОГО ФАЙЛА «ВЫГРУЗКИ» И НАРЕЗКА ====================

// Парсит готовый файл выгрузки полигона (например, Файл 5 и 6) и разбивает по станциям и датам
fun parseMasterVygruzkaFile(file: File): Map<String, Map<String, List<String>>> {
    val text = readDocumentTextAuto(file)
    if (text.isBlank()) return emptyMap()

    val result = mutableMapOf<String, MutableMap<String, MutableList<String>>>()
    var currentStation = ""
    var currentDateGroup = ""

    for (rawLine in text.lines()) {
        val line = rawLine.trimEnd()
        if (line.isBlank()) continue

        // Станция объявляется с начала строки без отступа
        if (!line.startsWith(" ") && !line.startsWith("\t")) {
            val stName = line.trim()
            if (stName.isNotEmpty() && !stName.contains("Изменения", true)) {
                currentStation = stName
                result.getOrPut(currentStation) { mutableMapOf() }
                currentDateGroup = ""
            }
            continue
        }

        val trimmed = line.trim()
        if (trimmed.startsWith("Изменений нет", true) || trimmed.contains("Изменения в графике", true) || trimmed.contains("на Ленинградский", true)) {
            continue
        }

        // Блок даты: "27 октября", "15, 23 октября" и т.д.
        if (Regex("^(\\d{1,2}(?:,\\s*\\d{1,2})*\\s+[а-я]+|с\\s*\\d+.*)").containsMatchIn(trimmed)) {
            currentDateGroup = trimmed
            if (currentStation.isNotEmpty()) {
                result.getOrPut(currentStation) { mutableMapOf() }.getOrPut(currentDateGroup) { mutableListOf() }
            }
            continue
        }

        // Строка поезда: "№6504 до Клина отправлением в..."
        if (currentStation.isNotEmpty() && currentDateGroup.isNotEmpty() && (trimmed.startsWith("№") || trimmed.startsWith("Назначается"))) {
            result.getOrPut(currentStation) { mutableMapOf() }.getOrPut(currentDateGroup) { mutableListOf() }.add(trimmed)
        }
    }

    return result
}

// Нарезка файлов Word (.docx) строго по утверждённому образцу
fun generateWordVygruzkaDocs(
    selectedStations: Set<String>,
    masterData: Map<String, Map<String, List<String>>>,
    selectedDates: Set<SingleDateFilterOption>,
    outputFolder: File
) {
    if (!outputFolder.exists()) outputFolder.mkdirs()

    for (targetStation in selectedStations) {
        val stKey = normalizeStationKey(targetStation)
        val stationEntry = masterData.entries.firstOrNull { normalizeStationKey(it.key) == stKey }
        val dateMap = stationEntry?.value ?: emptyMap()

        val docTo = XWPFDocument()
        val docFrom = XWPFDocument()

        // 1. Шапка документа "На Москву"
        val pHeaderTo1 = docTo.createParagraph().apply { alignment = ParagraphAlignment.CENTER }
        pHeaderTo1.createRun().apply {
            isBold = true
            fontSize = 14
            setText("$targetStation  Изменения в графике движения эл/поездов")
        }
        val pHeaderTo2 = docTo.createParagraph().apply { alignment = ParagraphAlignment.CENTER }
        pHeaderTo2.createRun().apply {
            isBold = true
            fontSize = 14
            setText("на Ленинградский вокзал, Ипподром")
        }

        // 2. Шапка документа "От Москвы"
        val pHeaderFrom1 = docFrom.createParagraph().apply { alignment = ParagraphAlignment.CENTER }
        pHeaderFrom1.createRun().apply {
            isBold = true
            fontSize = 14
            setText("$targetStation  Изменения в графике движения эл/поездов")
        }
        val pHeaderFrom2 = docFrom.createParagraph().apply { alignment = ParagraphAlignment.CENTER }
        pHeaderFrom2.createRun().apply {
            isBold = true
            fontSize = 14
            setText("от Москвы")
        }

        var hasTrainsTo = false
        var hasTrainsFrom = false

        for ((dateGroup, trainLines) in dateMap) {
            // Фильтрация по выбранным датам, если они заданы
            if (selectedDates.isNotEmpty()) {
                val matchesAnyDate = selectedDates.any { opt -> dateGroup.contains(opt.day.toString()) }
                if (!matchesAnyDate) continue
            }

            val toLines = mutableListOf<String>()
            val fromLines = mutableListOf<String>()

            for (tLine in trainLines) {
                val numMatch = Regex("\\b([67]\\d{3}(?:/[67]\\d{3})?)\\b").find(tLine)
                val primaryNum = numMatch?.groupValues?.get(1)?.split("/")?.first() ?: ""
                val dir = getDirectionBySingleNumber(primaryNum)
                if (dir == Direction.TO_MOSCOW) toLines.add(tLine) else fromLines.add(tLine)
            }

            if (toLines.isNotEmpty()) {
                hasTrainsTo = true
                val pDate = docTo.createParagraph().apply { alignment = ParagraphAlignment.CENTER }
                pDate.createRun().apply {
                    isBold = true
                    fontSize = 14
                    setText(dateGroup)
                }
                for (line in toLines) {
                    val pTrain = docTo.createParagraph()
                    pTrain.createRun().apply {
                        fontSize = 12
                        setText("        $line")
                    }
                }
            }

            if (fromLines.isNotEmpty()) {
                hasTrainsFrom = true
                val pDate = docFrom.createParagraph().apply { alignment = ParagraphAlignment.CENTER }
                pDate.createRun().apply {
                    isBold = true
                    fontSize = 14
                    setText(dateGroup)
                }
                for (line in fromLines) {
                    val pTrain = docFrom.createParagraph()
                    pTrain.createRun().apply {
                        fontSize = 12
                        setText("        $line")
                    }
                }
            }
        }

        // Сохраняем Word документы (.docx)
        if (hasTrainsTo) {
            val f = File(outputFolder, "$targetStation на Москву.docx")
            FileOutputStream(f).use { fos -> docTo.write(fos) }
        }
        if (hasTrainsFrom) {
            val f = File(outputFolder, "$targetStation от Москвы.docx")
            FileOutputStream(f).use { fos -> docFrom.write(fos) }
        }
    }
}

// ==================== 7. ПАРСИНГ И АУДИТ ЯНДЕКС (PDF) ====================

fun parseYandexPdfSchedule(file: File): Pair<String, List<YandexParsedTrain>> {
    val text = readDocumentTextAuto(file)
    if (text.isBlank()) return Pair("", emptyList())

    val stMatch = Regex("Расписание электричек по\\s*вокзалу\\s+([А-Яа-яЁё0-9\\s\\-]+)", RegexOption.IGNORE_CASE).find(text)
    val stClean = stMatch?.groupValues?.get(1)?.split("\n")?.firstOrNull()?.replace("Все направления", "")?.trim() ?: ""

    val list = mutableListOf<YandexParsedTrain>()
    val tokens = text.split(Regex("(?=\\b\\d{2}:\\d{2}\\b)"))

    for (t in tokens) {
        val cleanT = t.trim()
        if (cleanT.length < 5 || !cleanT.take(5).matches(Regex("\\d{2}:\\d{2}"))) continue
        val depTime = cleanT.take(5)
        val rest = cleanT.substring(5).trim()

        val numMatch = Regex("(?:[^\\d]|^)([67]\\d{3}(?:/[67]\\d{3})?)(?:[^\\d]|$)").find(rest) ?: continue
        val trainNum = numMatch.groupValues[1]

        val routePart = rest.substring(0, numMatch.range.first).replace("—", "–").replace("-", "–").trim()

        var stopsStr = ""
        if (rest.contains("Остановки:")) {
            stopsStr = rest.substringAfter("Остановки:").trim()
        } else if (rest.contains("Со всеми остановками", true) || rest.contains("Со всеми остановками", true)) {
            stopsStr = "Со всеми остановками"
        } else if (rest.contains("Без остановок", true)) {
            stopsStr = "Без остановок"
        }

        val nextTimeMatch = Regex("\\b\\d{2}:\\d{2}\\b").find(stopsStr)
        if (nextTimeMatch != null) {
            stopsStr = stopsStr.substring(0, nextTimeMatch.range.first).trim()
        }

        list.add(
            YandexParsedTrain(
                trainNum = trainNum,
                departureTime = depTime,
                route = routePart,
                stopsRule = stopsStr,
                rawText = cleanT.take(250)
            )
        )
    }

    return Pair(stClean, list)
}

fun auditYandexSchedule(
    yandexStation: String,
    yandexTrains: List<YandexParsedTrain>,
    afishaDb: Map<String, AfishaTrainInfo>,
    telegramEvents: List<TelegramTrainEvent>
): List<YandexDiscrepancyRow> {
    val normSt = normalizeStationKey(yandexStation)
    val discrepancies = mutableListOf<YandexDiscrepancyRow>()

    for (yTrain in yandexTrains) {
        val baseNums = splitDirectionalTrainNumbers(yTrain.trainNum)
        val primaryNum = baseNums.first()
        val dir = getDirectionBySingleNumber(primaryNum)

        val telEvents = telegramEvents.filter { ev ->
            splitDirectionalTrainNumbers(ev.trainNum).any { it in baseNums }
        }

        val afInfo = lookupAfisha(afishaDb, primaryNum)
        val afTime = afInfo?.stationTimes?.get(normSt) ?: ""

        val isCancelledInTel = telEvents.any { it.isCancelledFull }
        if (isCancelledInTel) {
            val doc = telEvents.firstOrNull { it.isCancelledFull }?.telegramHeader ?: "Телеграмма"
            discrepancies.add(
                YandexDiscrepancyRow(
                    trainNum = yTrain.trainNum,
                    route = yTrain.route,
                    direction = dir,
                    yandexTime = yTrain.departureTime,
                    actualExpectedTime = "ОТМЕНЕН",
                    yandexStops = yTrain.stopsRule,
                    telegramStopsChange = "Поезд полностью отменен",
                    basisDocument = doc,
                    discrepancyDescription = "КРИТИЧНО: Поезд отменен по телеграмме, но в Яндексе отображается как действующий!",
                    severity = YandexDiscrepancySeverity.CRITICAL_CANCEL_MISSED
                )
            )
            continue
        }

        var expectedTime = afTime
        var timeSource = "Афиша (базовый график)"
        val telWithTime = telEvents.firstOrNull { it.stationTimes.containsKey(normSt) }
        if (telWithTime != null) {
            expectedTime = telWithTime.stationTimes[normSt]?.first ?: expectedTime
            timeSource = telWithTime.telegramHeader
        }

        if (expectedTime.isNotEmpty() && expectedTime != "-" && expectedTime != yTrain.departureTime) {
            val diff = calculateMinutesDiffLabel(expectedTime, yTrain.departureTime)
            discrepancies.add(
                YandexDiscrepancyRow(
                    trainNum = yTrain.trainNum,
                    route = yTrain.route,
                    direction = dir,
                    yandexTime = yTrain.departureTime,
                    actualExpectedTime = expectedTime,
                    yandexStops = yTrain.stopsRule,
                    telegramStopsChange = "Сдвиг времени",
                    basisDocument = timeSource,
                    discrepancyDescription = "Несовпадение времени: в Яндексе ${yTrain.departureTime}, по РЖД должно быть $expectedTime $diff",
                    severity = YandexDiscrepancySeverity.TIME_MISMATCH
                )
            )
        }

        for (telEv in telEvents) {
            for (skipped in telEv.skippedStops) {
                val yLow = yTrain.stopsRule.lowercase()
                val isMentionedAsExcludedInYandex = yLow.contains("кроме") && yLow.contains(skipped.lowercase().take(5))
                val isExpressNoStops = yTrain.stopsRule.contains("Без остановок", true)

                if (!isMentionedAsExcludedInYandex && !isExpressNoStops) {
                    discrepancies.add(
                        YandexDiscrepancyRow(
                            trainNum = yTrain.trainNum,
                            route = yTrain.route,
                            direction = dir,
                            yandexTime = yTrain.departureTime,
                            actualExpectedTime = expectedTime.ifEmpty { yTrain.departureTime },
                            yandexStops = yTrain.stopsRule,
                            telegramStopsChange = "Отменена остановка: $skipped",
                            basisDocument = telEv.telegramHeader,
                            discrepancyDescription = "В телеграмме отменена остановка $skipped, а в расписании Яндекса эта отмена отсутствует!",
                            severity = YandexDiscrepancySeverity.STOP_MISMATCH
                        )
                    )
                }
            }
        }
    }

    return discrepancies
}

// ==================== 8. СРАВНЕНИЕ И ОБЪЕДИНЕНИЕ ИЗМЕНЕНИЙ ====================

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



// ==================== 9. ГЛАВНЫЙ ИНТЕРФЕЙС TRAIN-R ПО МАКЕТУ РИС. 1 И 2 ====================

@OptIn(androidx.compose.foundation.layout.ExperimentalLayoutApi::class)
@Composable
fun TrainRApp() {
    val cachedAfisha = remember { getCachedAfishaFile() }
    var afishaFile by remember { mutableStateOf<File?>(if (cachedAfisha.exists()) cachedAfisha else null) }
    var telegramFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var masterVygruzkaFile by remember { mutableStateOf<File?>(null) }
    var parsedMasterVygruzka by remember { mutableStateOf>>>(emptyMap()) }
    
    var yandexFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var yandexStationName by remember { mutableStateOf("") }
    var yandexTrainsList by remember { mutableStateOf<List<YandexParsedTrain>>(emptyList()) }
    var yandexAuditResults by remember { mutableStateOf<List<YandexDiscrepancyRow>>(emptyList()) }

    var afishaDb by remember { mutableStateOf>(emptyMap()) }
    var telegramEvents by remember { mutableStateOf<List<TelegramTrainEvent>>(emptyList()) }
    var globalErrors by remember { mutableStateOf<List<GlobalError>>(emptyList()) }
    var ignoredErrors by remember { mutableStateOf<Set<String>>(emptySet()) }
    var availableDates by remember { mutableStateOf<List<SingleDateFilterOption>>(emptyList()) }
    var selectedDates by remember { mutableStateOf<Set<SingleDateFilterOption>>(emptySet()) }
    
    var currentStation by remember { mutableStateOf("Химки") }
    var selectedStationsForExport by remember { mutableStateOf(setOf("Химки", "Зеленоград-Крюково", "Сходня", "Алабушево")) }
    
    var toMoscowList by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }
    var fromMoscowList by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }
    
    var activeNavTab by remember { mutableStateOf(MainNavTab.VYGRUZKI) }
    var activeDirectionTab by remember { mutableStateOf(0) }
    var searchQuery by remember { mutableStateOf("") }
    var globalSearchResults by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }
    var inspectedTrainRow by remember { mutableStateOf<StationComparisonRow?>(null) }
    
    var isBusy by remember { mutableStateOf(false) }
    var statusText by remember { mutableStateOf("Готов к работе. Загрузите файлы Афиши, Телеграмм или файл Выгрузки.") }
    var activeTrainFilter by remember { mutableStateOf(TrainFilter.ALL) }
    var showCalendarDialog by remember { mutableStateOf(false) }
    val scope = rememberCoroutineScope()
    var recalcJob by remember { mutableStateOf<Job?>(null) }

    LaunchedEffect(afishaFile) {
        if (afishaFile != null && afishaFile!!.exists()) {
            withContext(Dispatchers.IO) {
                val loaded = loadAfishaDatabase(afishaFile)
                afishaDb = loaded
                statusText = "Базовая Афиша активна (распознано ${loaded.size} ниток)."
            }
        }
    }

    fun openFileDialog(title: String, multi: Boolean): List<File> {
        val dlg = FileDialog(null as Frame?, title, FileDialog.LOAD).apply {
            isMultipleMode = multi
            isVisible = true
        }
        val res = dlg.files.toList()
        dlg.dispose()
        return res
    }

    fun recalculateStation(stName: String, dates: Set<SingleDateFilterOption> = selectedDates) {
        recalcJob?.cancel()
        recalcJob = scope.launch {
            isBusy = true
            delay(AppConstants.REPARSE_DEBOUNCE_MS)
            statusText = if (stName == "ВСЕ") "Сверка ВСЕХ станций..." else "Сверка расписания по ст. $stName..."
            val res = withContext(Dispatchers.IO) {
                val telsPair = TelegramParseCache.get(telegramFiles)
                val datesList = extractAvailableDateOptions(telsPair.first)
                if (stName == "ВСЕ") {
                    val allTo = mutableListOf<StationComparisonRow>()
                    val allFrom = mutableListOf<StationComparisonRow>()
                    POLYGON_STATIONS.forEach { st ->
                        val (t, f) = buildStationComparison(st, afishaDb, telsPair.first, emptyMap(), emptyMap(), dates, ignoredErrors)
                        allTo.addAll(t.map { it.copy(stationName = st) })
                        allFrom.addAll(f.map { it.copy(stationName = st) })
                    }
                    StationCalcResult(allTo.distinctBy { it.trainNum + it.direction.name + it.finalChangeText }, allFrom.distinctBy { it.trainNum + it.direction.name + it.finalChangeText }, datesList, telsPair.first, emptyMap(), telsPair.second)
                } else {
                    val (t, f) = buildStationComparison(stName, afishaDb, telsPair.first, emptyMap(), emptyMap(), dates, ignoredErrors)
                    StationCalcResult(t, f, datesList, telsPair.first, emptyMap(), telsPair.second)
                }
            }
            telegramEvents = res.telegramEvents
            availableDates = res.availableDates
            toMoscowList = res.toMoscow
            fromMoscowList = res.fromMoscow
            globalErrors = res.errors
            isBusy = false
            statusText = if (stName == "ВСЕ") "Сверка всех станций завершена." else "Станция $stName: На Москву — ${res.toMoscow.size}, От Москвы — ${res.fromMoscow.size}."
        }
    }

    MaterialTheme {
        Row(modifier = Modifier.fillMaxSize().background(TrainRTheme.BgMain)) {
            // ЛЕВЫЙ САЙДБАР (ТЁМНО-СИНИЙ ПО ДИЗАЙНУ РИСУНКА 2)
            Column(
                modifier = Modifier
                    .width(220.dp)
                    .fillMaxHeight()
                    .background(TrainRTheme.SidebarNavy)
                    .padding(16.dp),
                verticalArrangement = Arrangement.spacedBy(8.dp)
            ) {
                // ЛОГОТИП ВВЕРХУ
                Row(verticalAlignment = Alignment.CenterVertically) {
                    VintageTrainLogo(Modifier.size(36.dp))
                    Spacer(modifier = Modifier.width(8.dp))
                    Column {
                        Text("Train-R", color = Color.White, fontSize = 18.sp, fontWeight = FontWeight.Black)
                        Text("МТППК • МЦД-3", color = TrainRTheme.GoldAccent, fontSize = 9.sp, fontWeight = FontWeight.ExtraBold)
                    }
                }

                Spacer(modifier = Modifier.height(24.dp))

                // КНОПКА 1: ВЫГРУЗКИ (ОСНОВНОЕ ОКНО)
                SidebarNavButton(
                    title = "ВЫГРУЗКИ",
                    icon = Icons.Default.DateRange,
                    isSelected = activeNavTab == MainNavTab.VYGRUZKI,
                    onClick = { activeNavTab = MainNavTab.VYGRUZKI }
                )

                // КНОПКА 2: ОТЧЕТЫ / СВЕРКА
                SidebarNavButton(
                    title = "ОТЧЕТЫ / СВЕРКА",
                    icon = Icons.Default.Info,
                    isSelected = activeNavTab == MainNavTab.REPORTS_SVERKA,
                    onClick = { activeNavTab = MainNavTab.REPORTS_SVERKA }
                )

                // КНОПКА 3: РАСПИСАНИЕ / ЯНДЕКС
                SidebarNavButton(
                    title = "ЯНДЕКС / АФИША",
                    icon = Icons.Default.Place,
                    isSelected = activeNavTab == MainNavTab.SCHEDULE_YANDEX,
                    onClick = { activeNavTab = MainNavTab.SCHEDULE_YANDEX }
                )

                Spacer(modifier = Modifier.weight(1f))

                Surface(
                    shape = RoundedCornerShape(8.dp),
                    color = Color.White.copy(alpha = 0.08f),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Column(modifier = Modifier.padding(10.dp)) {
                        Text("Статус системы:", color = Color.White.copy(alpha = 0.7f), fontSize = 10.sp, fontWeight = FontWeight.Bold)
                        Text(statusText, color = Color.White, fontSize = 10.sp, maxLines = 4, overflow = TextOverflow.Ellipsis)
                    }
                }
            }

            // ПРАВАЯ ОСНОВНАЯ ОБЛАСТЬ
            Column(modifier = Modifier.weight(1f).fillMaxHeight().padding(14.dp)) {
                // ВЕРХНЯЯ СТРОКА: ПОИСК + 4 КНОПКИ ЗАГРУЗКИ (ЯНДЕКС, АФИША, ТЕЛЕГРАММА, ВЫГРУЗКА)
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg),
                    border = BorderStroke(1.dp, TrainRTheme.BorderCol)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(horizontal = 12.dp, vertical = 8.dp),
                        verticalAlignment = Alignment.CenterVertically,
                        horizontalArrangement = Arrangement.spacedBy(8.dp)
                    ) {
                        // Окно поиска поезда
                        OutlinedTextField(
                            value = searchQuery,
                            onValueChange = { searchQuery = it },
                            placeholder = { Text("Поиск поезда (№6504, 7529...)", fontSize = 12.sp) },
                            modifier = Modifier.width(260.dp).height(44.dp),
                            singleLine = true,
                            textStyle = LocalTextStyle.current.copy(fontSize = 12.sp, fontWeight = FontWeight.Bold),
                            shape = RoundedCornerShape(8.dp),
                            leadingIcon = { Icon(Icons.Default.Search, null, modifier = Modifier.size(16.dp)) },
                            trailingIcon = {
                                if (searchQuery.isNotEmpty()) {
                                    IconButton(onClick = { searchQuery = "" }) { Icon(Icons.Default.Clear, null, modifier = Modifier.size(14.dp)) }
                                }
                            }
                        )

                        Spacer(modifier = Modifier.weight(1f))

                        // КНОПКА ЯНДЕКС
                        TopBarLoadButton(
                            label = "ЯНДЕКС",
                            count = yandexFiles.size,
                            accent = TrainRTheme.YandexRed,
                            onClick = {
                                val f = openFileDialog("Выберите расписание Яндекс (.pdf)", false)
                                if (f.isNotEmpty()) {
                                    yandexFiles = f
                                    scope.launch {
                                        isBusy = true
                                        val (st, trains) = withContext(Dispatchers.IO) { parseYandexPdfSchedule(f.first()) }
                                        yandexStationName = st
                                        yandexTrainsList = trains
                                        val discr = auditYandexSchedule(st.ifEmpty { currentStation }, trains, afishaDb, telegramEvents)
                                        yandexAuditResults = discr
                                        isBusy = false
                                        statusText = "Яндекс расписание по ст. $st загружено (${trains.size} поездов)."
                                    }
                                }
                            }
                        )

                        // КНОПКА АФИША
                        TopBarLoadButton(
                            label = "АФИША",
                            count = if (afishaFile != null) afishaDb.size else 0,
                            accent = TrainRTheme.GreenSuccess,
                            onClick = {
                                val f = openFileDialog("Выберите файл базовой Афиши (.xlsx)", false)
                                if (f.isNotEmpty()) {
                                    try {
                                        FileInputStream(f.first()).use { inp -> FileOutputStream(cachedAfisha).use { out -> inp.copyTo(out) } }
                                        afishaFile = cachedAfisha
                                        afishaDb = loadAfishaDatabase(cachedAfisha)
                                    } catch (_: Exception) {
                                        afishaFile = f.first()
                                        afishaDb = loadAfishaDatabase(f.first())
                                    }
                                    if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates)
                                }
                            }
                        )

                        // КНОПКА ТЕЛЕГРАММА
                        TopBarLoadButton(
                            label = "ТЕЛЕГРАММА",
                            count = telegramFiles.size,
                            accent = TrainRTheme.RoyalBlue,
                            onClick = {
                                val f = openFileDialog("Выберите телеграммы (.docx / .pdf)", true)
                                if (f.isNotEmpty()) {
                                    telegramFiles = telegramFiles + f
                                    if (currentStation.isNotEmpty()) recalculateStation(currentStation, selectedDates)
                                }
                            }
                        )

                        // КНОПКА ВЫГРУЗКА (ФАЙЛ 5 И 6 ДЛЯ НАРЕЗКИ)
                        TopBarLoadButton(
                            label = "ВЫГРУЗКА",
                            count = if (masterVygruzkaFile != null) parsedMasterVygruzka.size else 0,
                            accent = TrainRTheme.PurpleMcd,
                            onClick = {
                                val f = openFileDialog("Выберите сводный файл выгрузки (.txt / .docx)", false)
                                if (f.isNotEmpty()) {
                                    masterVygruzkaFile = f.first()
                                    scope.launch {
                                        isBusy = true
                                        val parsed = withContext(Dispatchers.IO) { parseMasterVygruzkaFile(f.first()) }
                                        parsedMasterVygruzka = parsed
                                        isBusy = false
                                        statusText = "Файл выгрузки загружен! Распознано станций: ${parsed.size}."
                                    }
                                }
                            }
                        )
                    }
                }

                Spacer(modifier = Modifier.height(8.dp))

                // СТРОКА ДЕЙСТВИЙ (ФИЛЬТР ПО ДАТЕ, СВЕРИТЬ ВСЕ, СОЗДАТЬ ОТМЕНЫ)
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg),
                    border = BorderStroke(1.dp, TrainRTheme.BorderCol)
                ) {
                    Column(modifier = Modifier.fillMaxWidth().padding(10.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                            // ФИЛЬТР ПО ДАТЕ
                            Button(
                                onClick = { showCalendarDialog = true },
                                colors = ButtonDefaults.buttonColors(containerColor = Color(0xFFF1F5F9)),
                                border = BorderStroke(1.dp, TrainRTheme.BorderCol),
                                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp),
                                shape = RoundedCornerShape(6.dp)
                            ) {
                                Icon(Icons.Default.DateRange, null, tint = TrainRTheme.ActionBlue, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text(
                                    if (selectedDates.isEmpty()) "ФИЛЬТР ПО ДАТЕ: ВСЕ ДНИ" else "ВЫБРАНО ДАТ: ${selectedDates.size}",
                                    color = TrainRTheme.TextDark,
                                    fontSize = 11.sp,
                                    fontWeight = FontWeight.Black
                                )
                            }

                            Spacer(modifier = Modifier.weight(1f))

                            // СВЕРИТЬ ВСЕ
                            FilledTonalButton(
                                onClick = { recalculateStation("ВСЕ") },
                                enabled = !isBusy,
                                shape = RoundedCornerShape(6.dp)
                            ) {
                                Icon(Icons.Default.Refresh, null, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("СВЕРИТЬ ВСЕ", fontSize = 11.sp, fontWeight = FontWeight.Black)
                            }

                            // ГЛАВНАЯ СИНЯЯ КНОПКА: СОЗДАТЬ ОТМЕНЫ (НАРЕЗКА DOCX НА РАБОЧИЙ СТОЛ)
                            Button(
                                onClick = {
                                    scope.launch {
                                        isBusy = true
                                        statusText = "Формирование документов Word (.docx) по эталону на Рабочий стол..."
                                        withContext(Dispatchers.IO) {
                                            val desktop = File(System.getProperty("user.home"), "Desktop")
                                            val outFolder = File(desktop, "Отмены по станциям Word")
                                            generateWordVygruzkaDocs(
                                                selectedStations = selectedStationsForExport,
                                                masterData = parsedMasterVygruzka,
                                                selectedDates = selectedDates,
                                                outputFolder = outFolder
                                            )
                                        }
                                        isBusy = false
                                        statusText = "ГОТОВО! Файлы Word (.docx) по ${selectedStationsForExport.size} станциям созданы в папке «Отмены по станциям Word» на Рабочем столе."
                                    }
                                },
                                enabled = !isBusy && selectedStationsForExport.isNotEmpty(),
                                colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.ActionBlue),
                                contentPadding = PaddingValues(horizontal = 16.dp, vertical = 8.dp),
                                shape = RoundedCornerShape(6.dp)
                            ) {
                                Icon(Icons.Default.Check, null, modifier = Modifier.size(18.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("СОЗДАТЬ ОТМЕНЫ (${selectedStationsForExport.size} СТ.)", fontSize = 12.sp, fontWeight = FontWeight.Black)
                            }
                        }

                        Spacer(modifier = Modifier.height(8.dp))
                        Divider(color = TrainRTheme.BorderCol)
                        Spacer(modifier = Modifier.height(8.dp))

                        // СТРОКА СТАНЦИЙ ПОЛИГОНА (АЛАБУШЕВО, КРЮКОВО, ФИРСАНОВСКАЯ, СХОДНЯ, ХИМКИ...)
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text("СТАНЦИИ:", fontSize = 11.sp, fontWeight = FontWeight.Black, color = TrainRTheme.TextSub)
                            Spacer(modifier = Modifier.width(8.dp))
                            androidx.compose.foundation.layout.FlowRow(
                                modifier = Modifier.weight(1f),
                                horizontalArrangement = Arrangement.spacedBy(6.dp),
                                verticalArrangement = Arrangement.spacedBy(4.dp)
                            ) {
                                POLYGON_STATIONS.forEach { st ->
                                    val isChecked = st in selectedStationsForExport
                                    val isCurrent = currentStation == st
                                    Surface(
                                        shape = RoundedCornerShape(6.dp),
                                        color = if (isChecked) TrainRTheme.ActionBlue.copy(alpha = 0.12f) else Color(0xFFF1F5F9),
                                        border = BorderStroke(1.dp, if (isChecked) TrainRTheme.ActionBlue else TrainRTheme.BorderCol),
                                        modifier = Modifier.clickable {
                                            selectedStationsForExport = if (isChecked) selectedStationsForExport - st else selectedStationsForExport + st
                                            currentStation = st
                                            recalculateStation(st, selectedDates)
                                        }
                                    ) {
                                        Row(modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp), verticalAlignment = Alignment.CenterVertically) {
                                            Checkbox(
                                                checked = isChecked,
                                                onCheckedChange = { chk ->
                                                    selectedStationsForExport = if (chk) selectedStationsForExport + st else selectedStationsForExport - st
                                                },
                                                modifier = Modifier.size(16.dp)
                                            )
                                            Spacer(modifier = Modifier.width(4.dp))
                                            Text(
                                                st.replace("Зеленоград-", ""),
                                                fontSize = 11.sp,
                                                fontWeight = if (isCurrent) FontWeight.Black else FontWeight.Bold,
                                                color = if (isChecked) TrainRTheme.ActionBlue else TrainRTheme.TextDark
                                            )
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(8.dp))

                // ЦЕНТРАЛЬНАЯ РАБОЧАЯ ОБЛАСТЬ (В ЗАВИСИМОСТИ ОТ ВКЛАДКИ)
                Card(
                    modifier = Modifier.weight(1f).fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg),
                    border = BorderStroke(1.dp, TrainRTheme.BorderCol)
                ) {
                    when (activeNavTab) {
                        MainNavTab.VYGRUZKI -> {
                            Column(modifier = Modifier.fillMaxSize().padding(10.dp)) {
                                // Переключатель направления «НА МОСКВУ (НА ИППОДРОМ)» / «НА ОБЛАСТЬ (ИЗ ИППОДРОМА)»
                                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                                    TabRow(
                                        selectedTabIndex = activeDirectionTab,
                                        containerColor = TrainRTheme.CardBg,
                                        contentColor = TrainRTheme.RoyalBlue,
                                        modifier = Modifier.width(620.dp).height(38.dp).clip(RoundedCornerShape(8.dp))
                                    ) {
                                        Tab(
                                            selected = activeDirectionTab == 0,
                                            onClick = { activeDirectionTab = 0 },
                                            text = { Text("НА МОСКВУ (НА ИППОДРОМ) (${toMoscowList.size})", fontWeight = FontWeight.Black, fontSize = 11.sp) }
                                        )
                                        Tab(
                                            selected = activeDirectionTab == 1,
                                            onClick = { activeDirectionTab = 1 },
                                            text = { Text("НА ОБЛАСТЬ (ИЗ ИППОДРОМА) (${fromMoscowList.size})", fontWeight = FontWeight.Black, fontSize = 11.sp) }
                                        )
                                    }

                                    Row(horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                                        FilterChipItem("Все", activeTrainFilter == TrainFilter.ALL) { activeTrainFilter = TrainFilter.ALL }
                                        FilterChipItem("МЦД", activeTrainFilter == TrainFilter.MCD) { activeTrainFilter = TrainFilter.MCD }
                                        FilterChipItem("Экспрессы", activeTrainFilter == TrainFilter.EXPRESS) { activeTrainFilter = TrainFilter.EXPRESS }
                                    }
                                }

                                Spacer(modifier = Modifier.height(8.dp))

                                val baseList = if (activeDirectionTab == 0) toMoscowList else fromMoscowList
                                val filtered = if (searchQuery.isBlank()) baseList else baseList.filter { it.trainNum.contains(searchQuery.trim(), true) }

                                TrainTableSection(
                                    rows = filtered,
                                    emptyMessage = "Нет изменений по ст. $currentStation. Загрузите файлы и выберите даты.",
                                    onRowClick = { inspectedTrainRow = it }
                                )
                            }
                        }

                        MainNavTab.REPORTS_SVERKA -> {
                            StationAnalyticsReportView(
                                stationName = currentStation,
                                dateFilterLabel = selectedDates.joinToString(", ") { it.displayLabel }.ifEmpty { "Все даты" },
                                toMoscow = toMoscowList,
                                fromMoscow = fromMoscowList,
                                onTrainClick = { inspectedTrainRow = it }
                            )
                        }

                        MainNavTab.SCHEDULE_YANDEX -> {
                            YandexAuditTabView(
                                stationName = yandexStationName.ifEmpty { currentStation },
                                yandexTrains = yandexTrainsList,
                                discrepancies = yandexAuditResults,
                                onExportReports = {
                                    val target = yandexStationName.ifEmpty { currentStation }
                                    exportYandexAuditReports(target, yandexAuditResults)
                                    statusText = "Отчёты аудита Яндекса сохранены в папку на Рабочем столе!"
                                }
                            )
                        }
                    }
                }

                Spacer(modifier = Modifier.height(4.dp))

                if (isBusy) {
                    LinearProgressIndicator(modifier = Modifier.fillMaxWidth().height(3.dp).clip(RoundedCornerShape(2.dp)), color = TrainRTheme.ActionBlue, trackColor = Color(0xFFE2E8F0))
                } else {
                    Spacer(modifier = Modifier.height(3.dp))
                }
            }
        }

        if (showCalendarDialog) {
            CalendarGridPickerDialog(
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

        if (inspectedTrainRow != null) {
            TrainDossierDialog(row = inspectedTrainRow!!, onDismiss = { inspectedTrainRow = null })
        }
    }
}

// ==================== 10. ВСПОМОГАТЕЛЬНЫЕ UI КОМПОНЕНТЫ ====================

@Composable
fun SidebarNavButton(title: String, icon: androidx.compose.ui.graphics.vector.ImageVector, isSelected: Boolean, onClick: () -> Unit) {
    Surface(
        modifier = Modifier.fillMaxWidth().height(42.dp).clickable { onClick() },
        shape = RoundedCornerShape(8.dp),
        color = if (isSelected) TrainRTheme.SidebarActive else Color.Transparent
    ) {
        Row(modifier = Modifier.fillMaxSize().padding(horizontal = 12.dp), verticalAlignment = Alignment.CenterVertically) {
            Icon(icon, null, tint = Color.White, modifier = Modifier.size(18.dp))
            Spacer(modifier = Modifier.width(10.dp))
            Text(title, color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold)
        }
    }
}

@Composable
fun TopBarLoadButton(label: String, count: Int, accent: Color, onClick: () -> Unit) {
    Surface(
        modifier = Modifier.height(44.dp).clickable { onClick() },
        shape = RoundedCornerShape(8.dp),
        color = if (count > 0) accent.copy(alpha = 0.08f) else Color(0xFFF8FAFC),
        border = BorderStroke(1.dp, if (count > 0) accent else TrainRTheme.BorderCol)
    ) {
        Row(modifier = Modifier.padding(horizontal = 10.dp), verticalAlignment = Alignment.CenterVertically) {
            Icon(if (count > 0) Icons.Default.CheckCircle else Icons.Default.Add, null, tint = if (count > 0) accent else TrainRTheme.TextSub, modifier = Modifier.size(16.dp))
            Spacer(modifier = Modifier.width(6.dp))
            Column {
                Text(label, fontSize = 11.sp, fontWeight = FontWeight.Black, color = TrainRTheme.TextDark)
                Text(if (count > 0) "Активно: $count" else "Загрузить", fontSize = 9.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.Bold)
            }
        }
    }
}

@Composable
fun FilterChipItem(label: String, isSelected: Boolean, onClick: () -> Unit) {
    Surface(modifier = Modifier.height(26.dp).clickable { onClick() }, shape = RoundedCornerShape(13.dp), color = if (isSelected) TrainRTheme.RoyalBlue else Color(0xFFF1F5F9), border = BorderStroke(1.dp, if (isSelected) TrainRTheme.RoyalBlue else TrainRTheme.BorderCol)) {
        Box(contentAlignment = Alignment.Center, modifier = Modifier.padding(horizontal = 10.dp)) { Text(label, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = if (isSelected) Color.White else TrainRTheme.TextDark, maxLines = 1) }
    }
}

@Composable
fun TrainTableSection(rows: List<StationComparisonRow>, emptyMessage: String, showStationColumn: Boolean = false, onRowClick: (StationComparisonRow) -> Unit) {
    Column(modifier = Modifier.fillMaxSize()) {
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
                        Row(modifier = Modifier.width(95.dp), verticalAlignment = Alignment.CenterVertically) { Text("№${item.trainNum}", fontWeight = FontWeight.Black, fontFamily = FontFamily.Monospace, fontSize = 12.sp, color = TrainRTheme.NavyDark, maxLines = 1); if (item.isMcd) { Spacer(modifier = Modifier.width(2.dp)); Surface(shape = RoundedCornerShape(2.dp), color = Color(0xFFFFE4E6)) { Text("МЦД", fontSize = 8.sp, fontWeight = FontWeight.Black, color = Color(0xFFBE123C), modifier = Modifier.padding(horizontal = 2.dp)) } } }
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
    var searchInReport by remember { mutableStateOf("") }
    val all = (toMoscow + fromMoscow).distinctBy { "${it.trainNum}_${it.direction}_${it.dateGroup}" }
    
    val cancelledTrains = all.filter { it.isCancelledAtStation && !it.isPartialCancel }
    val partialCancelled = all.filter { it.isPartialCancel }
    val appointedTrains = all.filter { it.isAppointedAtStation }
    val appointedStops = all.filter { it.isAppointedStop }
    val shifted = all.filter { !it.isCancelledAtStation && !it.isAppointedAtStation && !it.isAppointedStop && (it.isTimeShiftedAtStation || it.finalChangeText.contains("отправится", true)) }
    val discrepancies = all.filter { it.hasDiscrepancyWithOtmeny }

    val reportScrollState = rememberScrollState()

    Column(modifier = Modifier.fillMaxSize().verticalScroll(reportScrollState).padding(16.dp)) {
        Row(
            modifier = Modifier.fillMaxWidth(),
            horizontalArrangement = Arrangement.SpaceBetween,
            verticalAlignment = Alignment.CenterVertically
        ) {
            Column {
                Text("Сводная отчетность по станции: $stationName", fontSize = 18.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                Text("Период выборки: $dateFilterLabel • Нажмите на карточку для детализации списка внизу", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
            }
            if (selectedCategory != 0) {
                FilledTonalButton(
                    onClick = { selectedCategory = 0 },
                    contentPadding = PaddingValues(horizontal = 10.dp, vertical = 4.dp),
                    modifier = Modifier.height(30.dp)
                ) {
                    Text("Сбросить фильтр", fontSize = 11.sp, fontWeight = FontWeight.ExtraBold)
                }
            }
        }

        Spacer(modifier = Modifier.height(10.dp))

        Row(horizontalArrangement = Arrangement.spacedBy(10.dp), modifier = Modifier.fillMaxWidth()) {
            StatSummaryBox("Отменено поездов", cancelledTrains.size, TrainRTheme.RedCancel, selectedCategory == 1, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 1) 0 else 1 }
            StatSummaryBox("Назначено поездов", appointedTrains.size, TrainRTheme.GreenSuccess, selectedCategory == 2, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 2) 0 else 2 }
            StatSummaryBox("Изменено время отпр.", shifted.size, TrainRTheme.GoldAccent, selectedCategory == 3, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 3) 0 else 3 }
            StatSummaryBox("Расхождений", discrepancies.size, TrainRTheme.NavyDark, selectedCategory == 4, Modifier.weight(1f)) { selectedCategory = if (selectedCategory == 4) 0 else 4 }
        }

        Spacer(modifier = Modifier.height(12.dp))

        Row(
            modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 6.dp, horizontal = 10.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
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
            else -> cancelledTrains + partialCancelled + appointedTrains + appointedStops + shifted + discrepancies
        }

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

@Composable
fun StatSummaryBox(label: String, count: Int, color: Color, isSelected: Boolean, modifier: Modifier, onClick: () -> Unit) {
    Surface(
        modifier = modifier.clickable { onClick() },
        shape = RoundedCornerShape(8.dp),
        color = if (isSelected) color.copy(alpha = 0.25f) else color.copy(alpha = 0.08f),
        border = BorderStroke(if (isSelected) 2.5.dp else 1.dp, if (isSelected) color else color.copy(alpha = 0.35f))
    ) {
        Row(modifier = Modifier.padding(horizontal = 12.dp, vertical = 10.dp), verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.SpaceBetween) {
            Column {
                Text(label, fontSize = 12.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
                Text(if (isSelected) "Фильтр активен" else "Нажмите", fontSize = 9.sp, fontWeight = FontWeight.Bold, color = color)
            }
            Text(count.toString(), fontSize = 22.sp, fontWeight = FontWeight.Black, color = color)
        }
    }
}

@Composable
fun DetailedReportTrainCard(r: StationComparisonRow, accent: Color, onClick: () -> Unit) {
    val diffText = if (!r.isCancelledAtStation && r.afishaTime != "-" && r.telegramTime != "-" && r.telegramTime != "ОТМ") calculateMinutesDiffLabel(r.afishaTime, r.telegramTime) else ""
    val timeComparisonStr = when { r.isCancelledAtStation -> "${r.afishaTime} ➔ ОТМЕНЕН"; r.isAppointedAtStation && r.afishaTime == "-" -> "Назначен ➔ ${r.telegramTime}"; else -> "${r.afishaTime} ➔ ${r.telegramTime} $diffText".trim() }
    Row(
        modifier = Modifier
            .fillMaxWidth()
            .clip(RoundedCornerShape(6.dp))
            .background(Color(0xFFF8FAFC))
            .border(1.dp, accent.copy(alpha = 0.35f), RoundedCornerShape(6.dp))
            .clickable { onClick() }
            .padding(vertical = 7.dp, horizontal = 10.dp),
        verticalAlignment = Alignment.CenterVertically
    ) {
        Text(r.dateGroup, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.RoyalBlue, modifier = Modifier.width(120.dp))
        Text("№${r.trainNum}", fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Black, fontSize = 13.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(95.dp))
        Column(modifier = Modifier.width(175.dp)) {
            Text(r.routeString, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark, maxLines = 1, overflow = TextOverflow.Ellipsis)
            Text(r.destinationLabel, fontWeight = FontWeight.SemiBold, fontSize = 9.sp, color = TrainRTheme.TextSub)
        }
        Surface(shape = RoundedCornerShape(5.dp), color = accent.copy(alpha = 0.12f), modifier = Modifier.width(185.dp).padding(end = 10.dp)) {
            Text(timeComparisonStr, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Black, fontSize = 11.sp, color = accent, modifier = Modifier.padding(horizontal = 6.dp, vertical = 3.dp))
        }
        Text(r.finalChangeText, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
    }
}

@Composable
fun YandexAuditTabView(
    stationName: String,
    yandexTrains: List<YandexParsedTrain>,
    discrepancies: List<YandexDiscrepancyRow>,
    onExportReports: () -> Unit
) {
    var searchFilter by remember { mutableStateOf("") }
    val criticalCount = discrepancies.count { it.severity == YandexDiscrepancySeverity.CRITICAL_CANCEL_MISSED }
    val timeCount = discrepancies.count { it.severity == YandexDiscrepancySeverity.TIME_MISMATCH }
    val stopCount = discrepancies.count { it.severity == YandexDiscrepancySeverity.STOP_MISMATCH }

    Column(modifier = Modifier.fillMaxSize().padding(12.dp)) {
        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
            Column {
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text("Аудит расписания Яндекс vs РЖД", fontSize = 18.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                    Spacer(modifier = Modifier.width(8.dp))
                    Surface(shape = RoundedCornerShape(4.dp), color = TrainRTheme.YandexRed) {
                        Text("СТАНЦИЯ: ${stationName.ifEmpty { "ЗЕЛЕНОГРАД-КРЮКОВО" }.uppercase()}", fontSize = 10.sp, fontWeight = FontWeight.Black, color = Color.White, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp))
                    }
                }
                Text("Загружено из Яндекс PDF: ${yandexTrains.size} поездов • Выявлено расхождений: ${discrepancies.size}", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
            }

            Button(
                onClick = onExportReports,
                enabled = discrepancies.isNotEmpty(),
                colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.YandexRed),
                shape = RoundedCornerShape(6.dp),
                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp)
            ) {
                Icon(Icons.Default.Share, null, modifier = Modifier.size(16.dp))
                Spacer(modifier = Modifier.width(6.dp))
                Text("Генерировать выгрузки по Яндексу", fontSize = 11.sp, fontWeight = FontWeight.Bold)
            }
        }

        Spacer(modifier = Modifier.height(10.dp))

        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
            StatSummaryBox("Отменены, но в Яндексе есть", criticalCount, TrainRTheme.RedCancel, false, Modifier.weight(1f)) {}
            StatSummaryBox("Не совпадает время", timeCount, TrainRTheme.GoldAccent, false, Modifier.weight(1f)) {}
            StatSummaryBox("Пропущена отмена остановки", stopCount, TrainRTheme.PurpleMcd, false, Modifier.weight(1f)) {}
            StatSummaryBox("Всего расхождений", discrepancies.size, TrainRTheme.NavyDark, false, Modifier.weight(1f)) {}
        }

        Spacer(modifier = Modifier.height(10.dp))

        Row(
            modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 6.dp, horizontal = 8.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            Text("№ ПОЕЗДА", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(90.dp))
            Text("МАРШРУТ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(160.dp))
            Text("ВРЕМЯ ЯНДЕКС", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(100.dp))
            Text("ФАКТ РЖД", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(100.dp))
            Text("ОПИСАНИЕ НЕСОВПАДЕНИЯ С ТЕЛЕГРАММАМИ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
            Text("ОСНОВАНИЕ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(130.dp))
        }

        Spacer(modifier = Modifier.height(4.dp))

        val filteredDiscrepancies = if (searchFilter.isBlank()) discrepancies else discrepancies.filter { it.trainNum.contains(searchFilter.trim(), true) }

        if (filteredDiscrepancies.isEmpty()) {
            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                Text(if (yandexTrains.isEmpty()) "Загрузите расписание Яндекс (кнопка «ЯНДЕКС» вверху)." else "Расхождений с телеграммами не обнаружено.", fontSize = 13.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub)
            }
        } else {
            LazyColumn(modifier = Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(4.dp)) {
                items(filteredDiscrepancies) { row ->
                    val badgeCol = when (row.severity) {
                        YandexDiscrepancySeverity.CRITICAL_CANCEL_MISSED -> TrainRTheme.RedCancel
                        YandexDiscrepancySeverity.TIME_MISMATCH -> TrainRTheme.GoldAccent
                        YandexDiscrepancySeverity.STOP_MISMATCH -> TrainRTheme.PurpleMcd
                        else -> TrainRTheme.NavyDark
                    }

                    Row(
                        modifier = Modifier
                            .fillMaxWidth()
                            .clip(RoundedCornerShape(6.dp))
                            .background(Color.White)
                            .border(1.dp, badgeCol.copy(alpha = 0.35f), RoundedCornerShape(6.dp))
                            .padding(vertical = 6.dp, horizontal = 8.dp),
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Text("№${row.trainNum}", fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Black, fontSize = 12.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(90.dp))
                        Text(row.route, fontWeight = FontWeight.Bold, fontSize = 10.sp, color = TrainRTheme.TextDark, modifier = Modifier.width(160.dp), maxLines = 1, overflow = TextOverflow.Ellipsis)
                        Surface(shape = RoundedCornerShape(4.dp), color = Color(0xFFF1F5F9), modifier = Modifier.width(90.dp)) {
                            Text(row.yandexTime, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Bold, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.padding(horizontal = 4.dp, vertical = 2.dp), textAlign = TextAlign.Center)
                        }
                        Spacer(modifier = Modifier.width(10.dp))
                        Surface(shape = RoundedCornerShape(4.dp), color = badgeCol.copy(alpha = 0.12f), modifier = Modifier.width(90.dp)) {
                            Text(row.actualExpectedTime, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.Black, fontSize = 11.sp, color = badgeCol, modifier = Modifier.padding(horizontal = 4.dp, vertical = 2.dp), textAlign = TextAlign.Center)
                        }
                        Spacer(modifier = Modifier.width(10.dp))
                        Column(modifier = Modifier.weight(1f)) {
                            Text(row.discrepancyDescription, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = badgeCol)
                            if (row.yandexStops.isNotEmpty()) {
                                Text("В Яндексе: ${row.yandexStops}", fontSize = 9.sp, color = TrainRTheme.TextSub, maxLines = 1, overflow = TextOverflow.Ellipsis)
                            }
                        }
                        Text(row.basisDocument, fontWeight = FontWeight.SemiBold, fontSize = 10.sp, color = TrainRTheme.RoyalBlue, modifier = Modifier.width(130.dp), maxLines = 2, overflow = TextOverflow.Ellipsis)
                    }
                }
            }
        }
    }
}

fun exportYandexAuditReports(station: String, auditList: List<YandexDiscrepancyRow>) {
    val desktop = File(System.getProperty("user.home"), "Desktop")
    val dirYandex = File(desktop, "Аудит Яндекс Расписаний").apply { if (!exists()) mkdirs() }
    val toMoscowList = auditList.filter { it.direction == Direction.TO_MOSCOW }
    val fromMoscowList = auditList.filter { it.direction == Direction.FROM_MOSCOW }

    File(dirYandex, "$station на Москву (Яндекс расхождения).txt").printWriter(Charsets.UTF_8).use { out ->
        out.println("$station    Расхождения Яндекс.Расписаний против телеграмм/афиши")
        out.println("Направление: На Ленинградский вокзал, Ипподром\n")
        toMoscowList.forEach { r -> out.println("№${r.trainNum} (${r.route}) в ${r.yandexTime} — ${r.discrepancyDescription} [Основание: ${r.basisDocument}]") }
    }
    File(dirYandex, "$station от Москвы (Яндекс расхождения).txt").printWriter(Charsets.UTF_8).use { out ->
        out.println("$station    Расхождения Яндекс.Расписаний против телеграмм/афиши")
        out.println("Направление: От Ленинградского вокзала, Ипподрома на область\n")
        fromMoscowList.forEach { r -> out.println("№${r.trainNum} (${r.route}) в ${r.yandexTime} — ${r.discrepancyDescription} [Основание: ${r.basisDocument}]") }
    }
}

@Composable
fun TrainDossierDialog(row: StationComparisonRow, onDismiss: () -> Unit) {
    val rawTextScrollState = rememberScrollState()
    var selectedStationInDossier by remember { mutableStateOf<String?>(row.stationName) }

    Dialog(onDismissRequest = onDismiss) {
        Card(
            modifier = Modifier.width(840.dp).heightIn(min = 520.dp, max = 740.dp),
            shape = RoundedCornerShape(12.dp),
            colors = CardDefaults.cardColors(containerColor = Color.White),
            elevation = CardDefaults.cardElevation(8.dp)
        ) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text("ПОЛНЫЙ ОТЧЕТ ПО ПОЕЗДУ №${row.trainNum} ${row.daysFlag}", fontSize = 16.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                            Spacer(modifier = Modifier.width(8.dp))
                            Surface(shape = RoundedCornerShape(4.dp), color = TrainRTheme.RoyalBlue) {
                                Text("СТАНЦИЯ: ${row.stationName.uppercase()}", fontSize = 10.sp, fontWeight = FontWeight.Black, color = Color.White, modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp))
                            }
                        }
                        Text("Маршрут: ${row.routeString} (${row.destinationLabel}) • Основание: ${row.telegramHeader}", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
                    }
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, contentDescription = null) }
                }

                Divider(modifier = Modifier.padding(vertical = 6.dp))

                Text("Период действия: ${row.dateGroup}", fontWeight = FontWeight.ExtraBold, fontSize = 12.sp)
                Text("Время по ст. ${row.stationName}: Афиша ${row.afishaTime} ➔ Телеграмма ${row.telegramTime}", fontWeight = FontWeight.Black, fontSize = 12.sp, color = TrainRTheme.RoyalBlue)
                Text("Итоговый статус: ${row.finalChangeText}", fontWeight = FontWeight.Black, fontSize = 12.sp, color = if (row.isCancelledAtStation) TrainRTheme.RedCancel else TrainRTheme.GreenSuccess)
                if (row.trackSummaryForStation.isNotEmpty()) Text("Изменение пути: ${row.trackSummaryForStation}", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.RoyalBlue)

                Spacer(modifier = Modifier.height(6.dp))

                Row(modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(4.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 4.dp, horizontal = 8.dp), verticalAlignment = Alignment.CenterVertically) {
                    Text("СТАНЦИЯ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(180.dp))
                    Text("ПО АФИШЕ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(100.dp))
                    Text("ПО ТЕЛЕГРАММЕ", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(140.dp))
                    Text("РАЗНИЦА / СТАТУС", fontWeight = FontWeight.Black, fontSize = 10.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
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

                        Row(
                            modifier = Modifier.fillMaxWidth().background(if (isSelectedInModal) TrainRTheme.GoldAccent.copy(alpha = 0.15f) else Color.Transparent).clickable { selectedStationInDossier = dispName }.padding(vertical = 4.dp, horizontal = 8.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text(dispName, fontWeight = if (isCurrentViewSt || isSelectedInModal) FontWeight.Black else FontWeight.Bold, fontSize = 11.sp, modifier = Modifier.width(180.dp))
                            Text(afT, fontFamily = FontFamily.Monospace, fontSize = 11.sp, modifier = Modifier.width(100.dp))
                            Text(telT, fontFamily = FontFamily.Monospace, fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, modifier = Modifier.width(140.dp))
                            Text(diffLabel.ifEmpty { "По графику" }, fontSize = 10.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
                        }
                        Divider(color = Color(0xFFE2E8F0))
                    }
                }

                Spacer(modifier = Modifier.height(6.dp))
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
fun CalendarGridPickerDialog(
    availableDates: List<SingleDateFilterOption>,
    initialSelected: Set<SingleDateFilterOption>,
    onDismiss: () -> Unit,
    onConfirm: (Set<SingleDateFilterOption>) -> Unit
) {
    val currentSelection = remember { mutableStateOf(initialSelected) }
    val cal = Calendar.getInstance()
    val baseMonth = availableDates.firstOrNull()?.monthNum ?: (cal.get(Calendar.MONTH) + 1)
    val baseYear = cal.get(Calendar.YEAR)
    val nextMonth = if (baseMonth == 12) 1 else baseMonth + 1
    val nextYear = if (baseMonth == 12) baseYear + 1 else baseYear
    val monthsToShow = listOf(Pair(baseYear, baseMonth), Pair(nextYear, nextMonth))

    Dialog(onDismissRequest = onDismiss) {
        Card(
            modifier = Modifier.width(720.dp).heightIn(max = 640.dp),
            shape = RoundedCornerShape(12.dp),
            colors = CardDefaults.cardColors(containerColor = Color.White),
            elevation = CardDefaults.cardElevation(8.dp)
        ) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column {
                        Text("Календарь изменений на 2 месяца", fontSize = 18.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                        Text("Активны только числа с изменениями по телеграммам", fontSize = 11.sp, color = TrainRTheme.TextSub, fontWeight = FontWeight.SemiBold)
                    }
                    IconButton(onClick = onDismiss) { Icon(Icons.Default.Close, null) }
                }

                Spacer(modifier = Modifier.height(10.dp))

                Row(modifier = Modifier.weight(1f).fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(16.dp)) {
                    monthsToShow.forEach { (year, month) ->
                        MonthGridCalendar(
                            year = year,
                            monthNum = month,
                            availableDates = availableDates,
                            selectedDates = currentSelection.value,
                            modifier = Modifier.weight(1f),
                            onDateToggle = { opt ->
                                currentSelection.value = if (opt in currentSelection.value) currentSelection.value - opt else currentSelection.value + opt
                            }
                        )
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    TextButton(onClick = { currentSelection.value = emptySet() }) {
                        Text("Очистить все дни", color = TrainRTheme.RedCancel, fontWeight = FontWeight.Bold)
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        TextButton(onClick = onDismiss) { Text("Отмена", color = TrainRTheme.TextSub) }
                        Button(
                            onClick = { onConfirm(currentSelection.value) },
                            colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RoyalBlue),
                            shape = RoundedCornerShape(8.dp)
                        ) {
                            Text("Применить фильтр (${currentSelection.value.size})", fontWeight = FontWeight.Bold)
                        }
                    }
                }
            }
        }
    }
}

@Composable
fun MonthGridCalendar(
    year: Int,
    monthNum: Int,
    availableDates: List<SingleDateFilterOption>,
    selectedDates: Set<SingleDateFilterOption>,
    modifier: Modifier = Modifier,
    onDateToggle: (SingleDateFilterOption) -> Unit
) {
    val monthName = MONTH_NAMES_NOMINATIVE[monthNum] ?: "Месяц $monthNum"
    val daysInMonth = YearMonth.of(year, monthNum).lengthOfMonth()
    val cal = Calendar.getInstance().apply { set(year, monthNum - 1, 1) }
    var firstDayOfWeek = cal.get(Calendar.DAY_OF_WEEK) - 1
    if (firstDayOfWeek == 0) firstDayOfWeek = 7
    val dayHeaders = listOf("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс")

    Column(modifier = modifier.clip(RoundedCornerShape(8.dp)).background(Color(0xFFF8FAFC)).border(1.dp, TrainRTheme.BorderCol, RoundedCornerShape(8.dp)).padding(8.dp)) {
        Text("$monthName $year", fontSize = 14.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark, modifier = Modifier.fillMaxWidth(), textAlign = TextAlign.Center)
        Spacer(modifier = Modifier.height(8.dp))

        Row(modifier = Modifier.fillMaxWidth()) {
            dayHeaders.forEach { h ->
                Text(h, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = if (h in listOf("Сб", "Вс")) TrainRTheme.RedCancel else TrainRTheme.TextSub, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
            }
        }
        Spacer(modifier = Modifier.height(4.dp))

        val totalCells = 42
        val cells = (1..totalCells).map { idx ->
            val dayNumber = idx - (firstDayOfWeek - 1)
            if (dayNumber in 1..daysInMonth) dayNumber else null
        }

        cells.chunked(7).forEach { week ->
            Row(modifier = Modifier.fillMaxWidth().padding(vertical = 2.dp)) {
                week.forEach { d ->
                    if (d == null) {
                        Spacer(modifier = Modifier.weight(1f).height(32.dp))
                    } else {
                        val option = availableDates.firstOrNull { it.monthNum == monthNum && it.day == d }
                        val isAvailable = option != null
                        val isSelected = option != null && option in selectedDates

                        Surface(
                            modifier = Modifier.weight(1f).height(32.dp).padding(1.dp).clickable(enabled = isAvailable) { if (option != null) onDateToggle(option) },
                            shape = RoundedCornerShape(6.dp),
                            color = if (isSelected) TrainRTheme.GoldAccent else if (isAvailable) Color(0xFFE2E8F0) else Color.Transparent,
                            border = if (isAvailable && !isSelected) BorderStroke(1.dp, TrainRTheme.RoyalBlue.copy(alpha = 0.5f)) else null
                        ) {
                            Box(contentAlignment = Alignment.Center) {
                                Text(d.toString(), fontSize = 11.sp, fontWeight = if (isAvailable) FontWeight.Black else FontWeight.Normal, color = if (isSelected) Color.White else if (isAvailable) TrainRTheme.NavyDark else Color(0xFFCBD5E1))
                            }
                        }
                    }
                }
            }
        }
    }
}

// ==================== 11. ТОЧКА ВХОДА (MAIN) ====================

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "Train-R — Интеллектуальный ревизор расписаний (Ленинградское направление • МЦД-3)",
        state = WindowState(width = 1380.dp, height = 880.dp)
    ) {
        TrainRApp()
    }
}
