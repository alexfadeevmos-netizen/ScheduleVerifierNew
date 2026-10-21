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
import androidx.compose.foundation.layout.FlowRow
import androidx.compose.foundation.layout.ExperimentalLayoutApi

// ==================== [NEW] КОНСТАНТЫ ====================

object AppConstants {
    const val DEFAULT_MONTH_NUM = 10
    const val DEFAULT_MONTH_NAME = "октября"
    const val UNGROUPED_DAY_KEY = 999
    const val DEFAULT_STATION_ORDER = 27 // Сходня по умолчанию
    const val TRACK_HINT_LOOKAHEAD = 45
    const val REPARSE_DEBOUNCE_MS = 150L
    const val SEARCH_DEBOUNCE_MS = 350L
    const val MIN_SEARCH_LEN = 4
}

// ==================== МОДЕЛИ ДАННЫХ TRAIN-R ====================

enum class Direction { TO_MOSCOW, FROM_MOSCOW }

data class AfishaTrainInfo(
    val trainNum: String,
    val daysFlag: String,
    val startStation: String,
    val endStation: String,
    val routeString: String,
    val destinationLabel: String,
    val isMcd: Boolean,
    val direction: Direction,
    val stationTimes: Map<String, String>,
    val stationTrackHints: Map<String, String>
)

data class TrackSegmentRule(
    val fromStationKey: String,
    val toStationKey: String,
    val trackNumber: String
)

data class CalendarDaySet(
    val monthNum: Int,
    val monthName: String,
    val days: List<Int>,
    val fallbackLabel: String
)

data class SingleDateFilterOption(
    val day: Int,
    val monthNum: Int,
    val monthName: String,
    val displayLabel: String
)

data class TelegramTrainEvent(
    val trainNum: String,
    val rawDates: String,
    val calendarDays: CalendarDaySet,
    val telegramHeader: String,
    val routeFromTelegram: String,
    val isCancelledFull: Boolean,
    val partialCancelSection: String,
    val isAppointedFull: Boolean,
    val stationTimes: LinkedHashMap<String, Pair<String, Boolean>>,
    val skippedStops: List<String>,
    val appointedStops: List<String>,
    val trackRules: List<TrackSegmentRule>,
    val explicitNoStopsText: String,
    val rawBlockText: String
)

data class StationComparisonRow(
    val dateGroup: String,
    val sortDayKey: Int,
    val activeDays: List<Int>,
    val monthName: String,
    val trainNum: String,
    val daysFlag: String,
    val routeString: String,
    val destinationLabel: String,
    val isMcd: Boolean,
    val direction: Direction,
    val stationName: String,
    val telegramHeader: String,
    val afishaTime: String,
    val telegramTime: String,
    val otmenyFileStatus: String,
    val finalChangeText: String,
    val isCancelledAtStation: Boolean,
    val isPassWithoutStop: Boolean,
    val isAppointedAtStation: Boolean,
    val isTimeShiftedAtStation: Boolean,
    val skippedStopsSummary: String,
    val appointedStopsSummary: String,
    val trackSummaryForStation: String,
    val hasDiscrepancyWithOtmeny: Boolean,
    val fullAfishaStations: Map<String, String>,
    val fullTelegramStations: Map<String, Pair<String, Boolean>>,
    val rawTelegramText: String
)

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

val CLOSED_OR_TRANSIT_STATIONS = setOf(
    "петровскоразумовская", "петрразумовская", "рижская", "пострижский", "малино"
)

fun getCachedAfishaFile(): File {
    val dir = File(System.getProperty("user.home"), ".train_r_system")
    if (!dir.exists()) dir.mkdirs()
    return File(dir, "cached_afisha.xlsx")
}

// ==================== [NEW] ЛОГИРОВАНИЕ ====================

private fun logWarn(msg: String, e: Throwable? = null) {
    System.err.println("[TrainR/WARN] $msg")
    if (e != null) e.printStackTrace()
}

// ==================== [NEW] ПАРНЫЕ НОМЕРА: 6712/6713 ====================

/**
 * [NEW] Возвращает направленные номера:
 *  - "6712"        -> ["6712"]
 *  - "6712/6713"   -> ["6712", "6713"]
 * Для парного МЦД-поезда каждый номер ведёт в свою сторону.
 */
fun splitDirectionalTrainNumbers(trainNum: String): List<String> =
    trainNum.split("/").map { it.trim() }.filter { it.isNotEmpty() }.distinct()

/**
 * [NEW] Ключи для поиска в справочниках, где номер может быть записан
 * как "6712", так и "6712/6713" (в любом порядке).
 */
fun trainNumLookupKeys(trainNum: String): Set<String> {
    val parts = splitDirectionalTrainNumbers(trainNum)
    val keys = parts.toMutableSet()
    if (parts.size >= 2) {
        keys.add("${parts[0]}/${parts[1]}")
        keys.add("${parts[1]}/${parts[0]}")
    }
    return keys
}

/** [NEW] Направление по ОДНОМУ номеру (чёт — от Москвы, нечёт — до Москвы). */
fun getDirectionBySingleNumber(singleNum: String): Direction {
    val digits = singleNum.filter { it.isDigit() }
    val lastDigit = digits.lastOrNull()?.digitToIntOrNull() ?: 1
    return if (lastDigit % 2 == 0) Direction.FROM_MOSCOW else Direction.TO_MOSCOW
}

/** @deprecated Используйте getDirectionBySingleNumber. Оставлено для совместимости. */
@Deprecated("Используйте getDirectionBySingleNumber")
fun getDirectionByTrainNumber(trainNum: String): Direction =
    getDirectionBySingleNumber(splitDirectionalTrainNumbers(trainNum).first())

/** [NEW] Поиск статуса в файле «Отмены» по направленному номеру с учётом парности. */
fun lookupOtmenyStatus(otmenyMap: Map<String, String>, directionalNum: String): String {
    otmenyMap[directionalNum]?.let { return it }
    return otmenyMap.entries.firstOrNull { (k, _) ->
        directionalNum in trainNumLookupKeys(k)
    }?.value ?: ""
}

/** [NEW] Поиск в Афише по направленному номеру с учётом парной записи. */
fun lookupAfisha(afishaDb: Map<String, AfishaTrainInfo>, directionalNum: String): AfishaTrainInfo? {
    afishaDb[directionalNum]?.let { return it }
    return afishaDb.entries.firstOrNull { (k, _) ->
        directionalNum in trainNumLookupKeys(k)
    }?.value
}

// ==================== ПРЕДОБРАБОТКА СЛИПШЕГОСЯ ТЕКСТА И СИНОНИМЫ СТАНЦИЙ ====================

fun preprocessTelegramLine(rawLine: String): String {
    var s = rawLine
    s = s.replace(Regex("(?i)(нр\\.?|поезд|№)(?=\\d)"), "$1 ")
    s = s.replace(Regex("(?<![:.])(\\b\\d{4}(?:/\\d{4})?)(?=[А-Яа-яЁё])"), "$1 ")
    s = s.replace(Regex("([А-Яа-яЁё])(?=\\d{1,2}[:.]\\d{2})"), "$1 ")
    s = s.replace(Regex("([А-Яа-яЁё0-9])\\("), "$1 (")
    return s.replace(Regex("\\s+"), " ").trim()
}

fun normalizeStationKey(name: String): String {
    val lower = name.lowercase().replace("ё", "е")
        .replace(Regex("\\(закрыта\\)"), "")
        .replace(Regex("\\(47\\s*км\\)"), "")
        .replace(Regex("\\(ст\\.?\\s*ховрино\\)"), "")
        .trim()

    if (Regex("\\b(оп|о\\.п\\.|пл|пл\\.)\\s*ховрин").containsMatchIn(lower)) {
        return "грачевская"
    }

    val clean = lower
        .replace(Regex("(?i)\\b(оп|о\\.п\\.|рп|р\\.п\\.|ст|ст\\.|пл|пл\\.|ост\\.?\\s*пункт)\\b"), "")
        .replace(Regex("[^а-я0-9]"), "")

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
        "крюково" -> "Крюково"
        "грачевская" -> "Грачёвская (ст. Ховрино)"
        "ховрино" -> "Ховрино"
        "новоподрезково" -> "Новоподрезково"
        "подрезково" -> "Подрезково"
        "молжаниново" -> "Молжаниново"
        "левобережная" -> "Левобережная"
        "фирсановская" -> "Фирсановская"
        "малино" -> "Малино"
        "сходня" -> "Сходня"
        "химки" -> "Химки"
        "моссельмаш" -> "Моссельмаш"
        "лихоборы" -> "Лихоборы"
        "останкино" -> "Останкино"
        "петровскоразумовская" -> "Петровско-Разумовская"
        "москватоварная" -> "Москва-Товарная"
        "москва" -> "Москва"
        "ипподром" -> "Ипподром"
        "алабушево" -> "Алабушево"
        "радищево" -> "Радищево"
        "поваровка" -> "Поваровка"
        "поварово1" -> "Поварово-1"
        "березкидачные" -> "Березки-Дачные"
        "подсолнечная" -> "Подсолнечная"
        "сенеж" -> "Сенеж"
        "головково" -> "Головково"
        "покровка" -> "Покровка"
        "фроловское" -> "Фроловское"
        "стреглово" -> "Стреглово"
        "клин" -> "Клин"
        "ямуга" -> "Ямуга"
        "решетниково" -> "Решетниково"
        "черничная" -> "Черничная"
        "завидово" -> "Завидово"
        "московскоеморе" -> "Московское Море"
        "редкино" -> "Редкино"
        "межево" -> "Межево"
        "кузьминка" -> "Кузьминка"
        "чуприяновка" -> "Чуприяновка"
        "тверь" -> "Тверь"
        "путепроводная" -> "Путепроводная"
        "конаковскиймох" -> "Конаковский Мох"
        "донховка" -> "Донховка"
        "конаковогрэс" -> "Конаково ГРЭС"
        else -> normKey.replaceFirstChar { it.uppercase() }
    }
}

fun cleanSingleTimeToken(raw: String): String {
    if (raw.isBlank()) return ""
    var s = raw.split("@")[0].trim()
    if ("/" in s) s = s.split("/").last().trim()
    s = s.replace(Regex("\\([+\\-]\\d+\\)"), "").trim()
    s = s.split(",")[0].trim().replace(".", ":")
    val nums = Regex("\\d+").findAll(s).map { it.value }.toList()
    if (nums.size >= 2) {
        val h = nums[0].toIntOrNull() ?: return ""
        val m = nums[1].toIntOrNull() ?: return ""
        if (h in 0..23 && m in 0..59) {
            return "%02d:%02d".format(h, m)
        }
    }
    return ""
}

fun calculateMinutesDiffLabel(afishaTime: String, telegramTime: String): String {
    val p1 = afishaTime.split(":")
    val p2 = telegramTime.split(":")
    if (p1.size < 2 || p2.size < 2) return ""
    val m1 = (p1[0].toIntOrNull() ?: return "") * 60 + (p1[1].toIntOrNull() ?: return "")
    val m2 = (p2[0].toIntOrNull() ?: return "") * 60 + (p2[1].toIntOrNull() ?: return "")
    var diff = m2 - m1
    if (diff < -720) diff += 1440
    if (diff > 720) diff -= 1440
    return when {
        diff > 0 -> "(+$diff мин)"
        diff < 0 -> "($diff мин)"
        else -> "(без сдвига)"
    }
}

// ==================== КАЛЕНДАРНЫЙ РАЗБОР ДАТ ====================

val MONTH_NAMES_BY_NUM = mapOf(
    1 to "января", 2 to "февраля", 3 to "марта", 4 to "апреля", 5 to "мая", 6 to "июня",
    7 to "июля", 8 to "августа", 9 to "сентября", 10 to "октября", 11 to "ноября", 12 to "декабря"
)

fun detectMonthNumFromString(s: String, defaultMonth: Int = AppConstants.DEFAULT_MONTH_NUM): Int {
    val low = s.lowercase()
    return when {
        "янв" in low -> 1; "фев" in low -> 2; "мар" in low -> 3; "апр" in low -> 4
        "мая" in low || "май" in low -> 5; "июн" in low -> 6; "июл" in low -> 7; "авг" in low -> 8
        "сен" in low -> 9; "окт" in low -> 10; "ноя" in low -> 11; "дек" in low -> 12
        else -> {
            val slashM = Regex("\\d{1,2}[/.\\-](\\d{1,2})").find(low)?.groupValues?.get(1)?.toIntOrNull()
            if (slashM != null && slashM in 1..12) slashM else defaultMonth
        }
    }
}

fun parseCalendarDays(raw: String, fallbackRaw: String): CalendarDaySet {
    val source = if (raw.isNotBlank()) raw else fallbackRaw
    val clean = source.replace("(", "").replace(")", "").trim()
    val mNum = detectMonthNumFromString(clean, detectMonthNumFromString(fallbackRaw, AppConstants.DEFAULT_MONTH_NUM))
    val mName = MONTH_NAMES_BY_NUM[mNum] ?: AppConstants.DEFAULT_MONTH_NAME

    val daysSet = sortedSetOf<Int>()

    val rangeRegex = Regex("(?i)с\\s*(\\d{1,2})(?:[/.\\-]\\d{1,2})?(?:\\s*[а-я]+)?\\s*по\\s*(\\d{1,2})(?:[/.\\-]\\d{1,2})?")
    for (rm in rangeRegex.findAll(clean)) {
        val d1 = rm.groupValues[1].toIntOrNull() ?: continue
        val d2 = rm.groupValues[2].toIntOrNull() ?: continue
        if (d1 in 1..31 && d2 in 1..31 && d1 <= d2) {
            for (d in d1..d2) daysSet.add(d)
        }
    }

    val listSlashRegex = Regex("((?:\\d{1,2}\\s*,\\s*)*\\d{1,2})[/.\\-](\\d{1,2})")
    for (lm in listSlashRegex.findAll(clean)) {
        val prefix = clean.substring(0, lm.range.first)
        if (prefix.trimEnd().endsWith("с", true) || prefix.trimEnd().endsWith("по", true) || prefix.trimEnd().endsWith("кроме", true) || prefix.trimEnd().endsWith("кр.", true)) {
            continue
        }
        lm.groupValues[1].split(",").forEach { p ->
            val d = p.trim().toIntOrNull()
            if (d != null && d in 1..31) daysSet.add(d)
        }
    }

    val onlyMatch = Regex("(?i)только\\s*(\\d{1,2})").find(clean)
    if (onlyMatch != null) {
        val d = onlyMatch.groupValues[1].toIntOrNull()
        if (d != null && d in 1..31) daysSet.add(d)
    }

    if (daysSet.isEmpty()) {
        val wordListMatch = Regex("((?:\\d{1,2}\\s*,\\s*)*\\d{1,2})\\s+[а-я]{3,}").find(clean)
        if (wordListMatch != null) {
            wordListMatch.groupValues[1].split(",").forEach { p ->
                val d = p.trim().toIntOrNull()
                if (d != null && d in 1..31) daysSet.add(d)
            }
        }
    }

    val exceptRegex = Regex("(?i)(?:кроме|кр\\.?)\\s*(\\d{1,2})")
    for (em in exceptRegex.findAll(clean)) {
        val exDay = em.groupValues[1].toIntOrNull()
        if (exDay != null) daysSet.remove(exDay)
    }

    return CalendarDaySet(
        monthNum = mNum,
        monthName = mName,
        days = daysSet.toList(),
        fallbackLabel = clean.ifEmpty { "По датам телеграммы" }
    )
}

fun formatDaysBlockHeader(days: List<Int>, monthName: String): String {
    if (days.isEmpty()) return "По датам телеграммы"
    val sorted = days.sorted()
    val isContiguous = (sorted.last() - sorted.first() + 1 == sorted.size)
    return when {
        sorted.size == 1 -> "${sorted.first()} $monthName"
        isContiguous && sorted.size >= 3 -> "с ${sorted.first()} по ${sorted.last()} $monthName"
        else -> "${sorted.joinToString(", ")} $monthName"
    }
}

fun extractAvailableDateOptions(events: List<TelegramTrainEvent>): List<SingleDateFilterOption> {
    val map = mutableMapOf<Pair<Int, Int>, SingleDateFilterOption>()
    for (ev in events) {
        val mNum = ev.calendarDays.monthNum
        val mName = ev.calendarDays.monthName
        for (d in ev.calendarDays.days) {
            val key = Pair(mNum, d)
            if (!map.containsKey(key)) {
                map[key] = SingleDateFilterOption(
                    day = d,
                    monthNum = mNum,
                    monthName = mName,
                    displayLabel = "$d $mName"
                )
            }
        }
    }
    return map.values.sortedWith(compareBy({ it.monthNum }, { it.day }))
}

// ==================== ОПРЕДЕЛЕНИЕ ПУТИ ДЛЯ СХОДНИ И ХИМОК ====================

fun formatTrackBadgeForStation(stationNorm: String, trackNum: String, rawHint: String = ""): String {
    if (stationNorm != "сходня" && stationNorm != "химки") return ""
    if (rawHint.isNotEmpty()) {
        val cleanHint = rawHint.removePrefix("@").trim()
        return "(по $cleanHint)"
    }
    return when (trackNum) {
        "1" -> "(по 1 пл. 1 пути)"
        "2" -> "(по 2 пл. 2 пути)"
        "3" -> if (stationNorm == "сходня") "(по 3 пл. 3 пути)" else "(по 3 пути)"
        "4" -> if (stationNorm == "химки") "(по 3 пл. 4 пути)" else "(по 4 пути)"
        else -> "(по $trackNum пути)"
    }
}

fun resolveTrackForStation(
    stationNorm: String,
    trainNum: String,
    direction: Direction,
    rules: List<TrackSegmentRule>,
    afishaHint: String
): String {
    if (stationNorm != "сходня" && stationNorm != "химки") return ""
    if (afishaHint.isNotEmpty()) {
        return formatTrackBadgeForStation(stationNorm, "", afishaHint)
    }
    val stOrder = STATION_ORDER_INDEX[stationNorm] ?: return ""

    for (rule in rules) {
        val idx1 = STATION_ORDER_INDEX[rule.fromStationKey] ?: continue
        val idx2 = STATION_ORDER_INDEX[rule.toStationKey] ?: continue
        val minIdx = minOf(idx1, idx2)
        val maxIdx = maxOf(idx1, idx2)

        if (stOrder in minIdx..maxIdx) {
            val isMovingTowardsMoscow = (direction == Direction.TO_MOSCOW)
            val isEndBoundary = (isMovingTowardsMoscow && stOrder == maxIdx && minIdx < maxIdx) ||
                    (!isMovingTowardsMoscow && stOrder == minIdx && minIdx < maxIdx)
            if (isEndBoundary && rules.size > 1) continue
            return formatTrackBadgeForStation(stationNorm, rule.trackNumber)
        }
    }
    return ""
}

// ==================== ПАРСЕР БАЗОВОЙ АФИШИ EXCEL ====================

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
                    val row = sheet.getRow(r) ?: continue
                    var foundTrains = 0
                    for (c in 1..minOf(50, row.lastCellNum.toInt())) {
                        val v = row.getCell(c)?.toString()?.trim()?.replace(".0", "") ?: ""
                        if (Regex("\\b\\d{4}\\b").containsMatchIn(v)) foundTrains++
                    }
                    if (foundTrains >= 3) {
                        headerRowIdx = r
                        break
                    }
                }

                val stationRows = mutableMapOf<Int, String>()
                for (r in (headerRowIdx + 1)..sheet.lastRowNum) {
                    val cellVal = sheet.getRow(r)?.getCell(0)?.toString()?.trim() ?: ""
                    if (cellVal.isNotEmpty() &&
                        !cellVal.contains("Условные", true) &&
                        !cellVal.contains("Расписание", true) &&
                        !cellVal.equals("изм", true) &&
                        !cellVal.equals("отм", true)
                    ) {
                        stationRows[r] = cellVal
                    }
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
                            trackHintsDb.getOrPut(trainNum) { mutableMapOf() }[normSt] = hint
                        }
                    }

                    if (db.containsKey(trainNum) && !isNormColumn) continue

                    val daysFlag = when {
                        combined.contains("СВ", true) -> "СВ"
                        combined.contains("Р", true) -> "Р"
                        else -> ""
                    }
