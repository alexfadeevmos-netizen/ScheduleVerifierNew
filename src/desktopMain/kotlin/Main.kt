import androidx.compose.foundation.Canvas
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
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
    val stationTimes: Map<String, String>
)

data class TelegramTrainEvent(
    val trainNum: String,
    val rawDates: String,
    val formattedDateGroup: String,
    val telegramHeader: String,
    val routeFromTelegram: String,
    val isCancelledFull: Boolean,
    val partialCancelSection: String,
    val isAppointedFull: Boolean,
    val stationTimes: Map<String, Pair<String, Boolean>>, // stationKey -> (depTime, hasStop)
    val skippedStops: List<String>,
    val appointedStops: List<String>,
    val trackNotes: String,
    val rawBlockText: String
)

data class StationComparisonRow(
    val dateGroup: String,
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
    val skippedStopsSummary: String,
    val appointedStopsSummary: String,
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
    valMcdCoral = Color(0xFFBE123C)
    val GoldAccent = Color(0xFFD97706)
    val GreenSuccess = Color(0xFF059669)
    val RedCancel = Color(0xFFDC2626)
    val AmberWarn = Color(0xFFB45309)
    val BorderCol = Color(0xFFCBD5E1)
    val TextDark = Color(0xFF0F172A)
    val TextSub = Color(0xFF334155)
}

// Станции участка выгрузки и всего полигона Москва - Тверь (Конаково) / МЦД-3
val POLYGON_STATIONS = listOf(
    "Сходня", "Химки", "Зеленоград-Крюково", "Подрезково", "Новоподрезково",
    "Молжаниново", "Левобережная", "Алабушево", "Фирсановская", "Ховрино",
    "Грачёвская", "Моссельмаш", "Лихоборы", "Останкино", "Поварово-1",
    "Подсолнечная", "Клин", "Решетниково", "Завидово", "Редкино", "Тверь", "Конаково ГРЭС"
)

// Закрытые или транзитные пункты, которые игнорируются при авто-проверке безостановочного проследования
val CLOSED_OR_TRANSIT_STATIONS = setOf(
    "петровскоразумовская", "петрразумовская", "рижская", "пострижский", "малино"
)

fun getCachedAfishaFile(): File {
    val dir = File(System.getProperty("user.home"), ".train_r_system")
    if (!dir.exists()) dir.mkdirs()
    return File(dir, "cached_afisha.xlsx")
}

// ==================== НОРМАЛИЗАЦИЯ СТАНЦИЙ И ВРЕМЕНИ ====================

fun normalizeStationKey(name: String): String {
    val clean = name.lowercase()
        .replace("ё", "е")
        .replace(Regex("\\(закрыта\\)"), "")
        .replace(Regex("(?i)\\b(оп|о\\.п\\.|рп|р\\.п\\.|ст|ст\\.|пл|пл\\.)\\b"), "")
        .replace(Regex("[^а-я0-9]"), "")
    return when {
        "крюков" in clean -> "крюково"
        "грачевск" in clean || "опховрино" in clean -> "грачевская"
        "ховрин" in clean -> "ховрино"
        "новоподрезков" in clean -> "новоподрезково"
        "подрезков" in clean -> "подрезково"
        "молжанинов" in clean -> "молжаниново"
        "левобереж" in clean -> "левобережная"
        "фирсанов" in clean -> "фирсановская"
        "сходн" in clean -> "сходня"
        "химк" in clean -> "химки"
        "моссельмаш" in clean -> "моссельмаш"
        "лихобор" in clean -> "лихоборы"
        "останкин" in clean -> "останкино"
        "петровско" in clean || "петрразум" in clean -> "петровскоразумовская"
        "постриж" in clean -> "пострижский"
        "рижск" in clean -> "рижская"
        "моск" in clean && "тов" in clean -> "москватоварная"
        "моск" in clean -> "москва"
        "ипподром" in clean || "47км" in clean -> "ипподром"
        "алабушев" in clean -> "алабушево"
        "радищев" in clean -> "радищево"
        "поваровк" in clean -> "поваровка"
        "поварово" in clean -> "поварово1"
        "березки" in clean -> "березкидачные"
        "подсолнеч" in clean -> "подсолнечная"
        "сенеж" in clean -> "сенеж"
        "головков" in clean -> "головково"
        "покровк" in clean -> "покровка"
        "фроловск" in clean -> "фроловское"
        "стреглов" in clean || "стриглов" in clean -> "стреглово"
        "клин" in clean -> "клин"
        "ямуг" in clean -> "ямуга"
        "решетников" in clean -> "решетниково"
        "завидов" in clean -> "завидово"
        "редкин" in clean -> "редкино"
        "твер" in clean -> "тверь"
        "конаков" in clean && "мох" in clean -> "конаковскиймох"
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
        "сходня" -> "Сходня"
        "химки" -> "Химки"
        "моссельмаш" -> "Моссельмаш"
        "лихоборы" -> "Лихоборы"
        "останкино" -> "Останкино"
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
        "решетниково" -> "Решетниково"
        "завидово" -> "Завидово"
        "редкино" -> "Редкино"
        "тверь" -> "Тверь"
        "конаковогрэс" -> "Конаково ГРЭС"
        else -> normKey.replaceFirstChar { it.uppercase() }
    }
}

// Правило пользователя: из "12.01:12:02", "14.55-14.55,5", "04:20/04:37" берем ТОЛЬКО ВРЕМЯ ОТПРАВЛЕНИЯ (второе)
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

// Определение направления строго по четности основного номера поезда
fun getDirectionByTrainNumber(trainNum: String): Direction {
    val firstPart = trainNum.split("/").first().filter { it.isDigit() }
    val lastDigit = firstPart.lastOrNull()?.digitToIntOrNull() ?: 1
    return if (lastDigit % 2 == 0) Direction.FROM_MOSCOW else Direction.TO_MOSCOW
}

// Форматирование заголовков дат по порядку (без дробления единых диапазонов)
fun formatTelegramDateGroup(raw: String, fallback: String): String {
    val clean = raw.replace("(", "").replace(")", "").trim()
    if (clean.isEmpty()) return fallback

    val months = mapOf(
        "01" to "января", "02" to "февраля", "03" to "марта", "04" to "апреля", "05" to "мая", "06" to "июня",
        "07" to "июля", "08" to "августа", "09" to "сентября", "10" to "октября", "11" to "ноября", "12" to "декабря",
        "1" to "января", "2" to "февраля", "3" to "марта", "4" to "апреля", "5" to "мая", "6" to "июня",
        "7" to "июля", "8" to "августа", "9" to "сентября"
    )

    // Если несколько диапазонов: "с 12/10 по 15/10, с 19/10 по 23/10"
    val multiRanges = Regex("с\\s*(\\d{1,2})[/.\\-](\\d{1,2})\\s*по\\s*(\\d{1,2})[/.\\-](\\d{1,2})").findAll(clean).toList()
    if (multiRanges.isNotEmpty()) {
        val parts = multiRanges.map { m ->
            val (d1, _, d2, m2) = m.destructured
            val mName = months[m2] ?: "октября"
            "с ${d1.toInt()} по ${d2.toInt()} $mName"
        }
        var res = parts.joinToString(", ")
        val kr = Regex("кроме\\s*(\\d{1,2})[/.\\-](\\d{1,2})").find(clean) ?: Regex("кр\\.?\\s*(\\d{1,2})[/.\\-](\\d{1,2})").find(clean)
        if (kr != null) {
            val mName = months[kr.groupValues[2]] ?: "октября"
            res += " (кроме ${kr.groupValues[1].toInt()} $mName)"
        }
        return res
    }

    val onlySingle = Regex("только\\s*(\\d{1,2})[/.\\-](\\d{1,2})").find(clean)
    if (onlySingle != null) {
        val mName = months[onlySingle.groupValues[2]] ?: "октября"
        return "${onlySingle.groupValues[1].toInt()} $mName"
    }

    val listMatch = Regex("([\\d\\s,]+)[/.\\-](\\d{1,2})").find(clean)
    if (listMatch != null) {
        val days = listMatch.groupValues[1].split(",").mapNotNull { it.trim().toIntOrNull() }
        if (days.isNotEmpty()) {
            val mName = months[listMatch.groupValues[2]] ?: "октября"
            return "${days.joinToString(", ")} $mName"
        }
    }

    return clean
}

fun extractFirstDayForSort(dateGroup: String): Int {
    return Regex("\\d+").find(dateGroup)?.value?.toIntOrNull() ?: 999
}

// ==================== ПАРСЕР БАЗОВОЙ АФИШИ EXCEL ====================

fun loadAfishaDatabase(afishaFile: File?): Map<String, AfishaTrainInfo> {
    if (afishaFile == null || !afishaFile.exists()) return emptyMap()
    val db = mutableMapOf<String, AfishaTrainInfo>()
    val timeFmt = SimpleDateFormat("HH:mm")

    try {
        FileInputStream(afishaFile).use { fis ->
            val wb = WorkbookFactory.create(fis)
            for (sIdx in 0 until wb.numberOfSheets) {
                val sheet = wb.getSheetAt(sIdx)
                val stationRows = mutableMapOf<Int, String>()
                for (r in 5..sheet.lastRowNum) {
                    val cellVal = sheet.getRow(r)?.getCell(0)?.toString()?.trim() ?: ""
                    if (cellVal.isNotEmpty() && !cellVal.contains("Условные", true) && !cellVal.contains("Расписание", true)) {
                        stationRows[r] = cellVal
                    }
                }

                // Ищем строку с номерами поездов (обычно строка 4, индекс 3)
                var headerRowIdx = 3
                for (r in 0..8) {
                    val row = sheet.getRow(r) ?: continue
                    var foundTrains = 0
                    for (c in 1..minOf(40, row.lastCellNum.toInt())) {
                        val v = row.getCell(c)?.toString()?.trim() ?: ""
                        if (Regex("\\b\\d{4}\\b").containsMatchIn(v)) foundTrains++
                    }
                    if (foundTrains >= 3) {
                        headerRowIdx = r
                        break
                    }
                }

                val hRow = sheet.getRow(headerRowIdx) ?: continue
                val subRow = sheet.getRow(headerRowIdx + 1)

                for (c in 1..hRow.lastCellNum) {
                    val v1 = hRow.getCell(c)?.toString()?.trim()?.replace(".0", "") ?: ""
                    val v2 = subRow?.getCell(c)?.toString()?.trim() ?: ""
                    val combined = "$v1 $v2".trim()
                    val numMatch = Regex("\\b(\\d{4})\\b").find(combined) ?: continue
                    val trainNum = numMatch.groupValues[1]

                    val isNormColumn = combined.contains("норм", true) || (!combined.contains("изм", true) && !combined.contains("отм", true))
                    if (db.containsKey(trainNum) && !isNormColumn) {
                        continue // Всегда сохраняем нормативный столбец ("норм") как базу Афиши!
                    }

                    val daysFlag = when {
                        combined.contains("СВ", true) -> "СВ"
                        combined.contains("Р", true) -> "Р"
                        else -> ""
                    }

                    val stTimes = LinkedHashMap<String, String>()
                    var firstStRaw = ""
                    var lastStRaw = ""

                    for ((rIdx, rawStName) in stationRows) {
                        val cell = sheet.getRow(rIdx)?.getCell(c) ?: continue
                        val rawTimeStr = if (cell.cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                            timeFmt.format(cell.dateCellValue)
                        } else {
                            cell.toString().trim()
                        }
                        if (rawTimeStr.contains("Конаково", true)) continue
                        val cleaned = cleanSingleTimeToken(rawTimeStr)
                        if (cleaned.isNotEmpty() && !rawStName.contains("закрыта", true)) {
                            val normKey = normalizeStationKey(rawStName)
                            stTimes[normKey] = cleaned
                            if (firstStRaw.isEmpty()) firstStRaw = rawStName
                            lastStRaw = rawStName
                        }
                    }

                    if (stTimes.isNotEmpty()) {
                        val startNorm = normalizeStationKey(firstStRaw)
                        val endNorm = normalizeStationKey(lastStRaw)
                        val isMcd = startNorm == "пострижский" || endNorm == "пострижский" || trainNum.startsWith("72") || trainNum.startsWith("73") || trainNum.startsWith("74")

                        val cleanStart = when (startNorm) {
                            "пострижский", "ипподром" -> "Ипподром"
                            "москва" -> "Москва"
                            else -> displayStationName(startNorm)
                        }
                        val cleanEnd = when (endNorm) {
                            "пострижский", "ипподром" -> "Ипподром"
                            "москва" -> "Москва"
                            else -> displayStationName(endNorm)
                        }

                        val dir = getDirectionByTrainNumber(trainNum)
                        val destLabel = if (dir == Direction.TO_MOSCOW) {
                            if (isMcd || cleanEnd == "Ипподром") "до Ипподрома (47 км)" else "до Москвы"
                        } else {
                            when (endNorm) {
                                "крюково" -> "до Крюково"
                                "клин" -> "до Клина"
                                "тверь" -> "до Твери"
                                "конаковогрэс" -> "до Конаково ГРЭС"
                                "подсолнечная" -> "до Подсолнечной"
                                "решетниково" -> "до Решетниково"
                                "поварово1" -> "до Поварово-1"
                                "поваровка" -> "до Поваровки"
                                else -> "до $cleanEnd"
                            }
                        }

                        db[trainNum] = AfishaTrainInfo(
                            trainNum = trainNum,
                            daysFlag = daysFlag,
                            startStation = cleanStart,
                            endStation = cleanEnd,
                            routeString = "$cleanStart – $cleanEnd",
                            destinationLabel = destLabel,
                            isMcd = isMcd,
                            direction = dir,
                            stationTimes = stTimes
                        )
                    }
                }
            }
        }
    } catch (_: Exception) {}
    return db
}

// ==================== ИНТЕЛЛЕКТУАЛЬНЫЙ ПАРСЕР ТЕЛЕГРАММ ====================

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
    } catch (_: Exception) { "" }
}

fun parseTelegramDocument(file: File): List<TelegramTrainEvent> {
    val text = readDocumentText(file)
    if (text.isBlank()) return emptyList()

    val headerMatch = Regex("(?i)ТЕЛЕГРАММА\\s*№?\\s*([^\\s]+)\\s+от\\s+([0-9/.\\-]+(?:г\\.?)?)").find(text)
    val telHeader = if (headerMatch != null) {
        "Тел. №${headerMatch.groupValues[1]} от ${headerMatch.groupValues[2]}"
    } else {
        "Тел. (${file.nameWithoutExtension.take(18)})"
    }

    val globalDatesMatch = Regex("(?i)(?:с\\s*\\d{1,2}\\s+[а-я]+\\s+по\\s+\\d{1,2}\\s+[а-я]+(?:,\\s*с\\s*\\d{1,2}\\s+[а-я]+\\s+по\\s+\\d{1,2}\\s+[а-я]+)?)").find(text)
    val defaultDateGroup = globalDatesMatch?.value?.trim() ?: "По датам телеграммы"

    val events = mutableListOf<TelegramTrainEvent>()
    val lines = text.lines()

    var inCancelBlock = false
    var inAppointBlock = false
    var inBologoeSection = false
    var currentSectionDates = ""

    var i = 0
    while (i < lines.size) {
        val line = lines[i].trim()
        if (line.isEmpty()) { i++; continue }

        // Отсекаем чужой участок Бологое - Окуловка - Торжок
        if (line.contains("Бологое", true) && (line.contains("участке", true) || line.contains("Окуловка", true))) {
            inBologoeSection = true
        } else if (line.contains("пригородных поездов МЦД", true) || line.contains("пригородных поездов:", true) || line.contains("«Ласточка»", true)) {
            inBologoeSection = false
        }

        // Определяем режим секции (Отменить / Назначить / Изменить)
        if (Regex("(?i)(?:^|\\s|[0-9.]+)(?:отменить|не назначать)").containsMatchIn(line) && (line.contains(":") || line.contains("("))) {
            inCancelBlock = true
            inAppointBlock = false
            val secDate = Regex("\\(([^)]+)\\)").find(line)?.groupValues?.get(1)
            if (secDate != null) currentSectionDates = secDate
            i++
            continue
        } else if (Regex("(?i)(?:^|\\s|[0-9.]+)назначить\\s*:").containsMatchIn(line)) {
            inCancelBlock = false
            inAppointBlock = true
            currentSectionDates = ""
            i++
            continue
        } else if (Regex("(?i)(?:^|\\s|[0-9.]+)изменить(?:\\s+расписание)?\\s*:").containsMatchIn(line)) {
            inCancelBlock = false
            inAppointBlock = false
            val secDate = Regex("\\(([^)]+)\\)").find(line)?.groupValues?.get(1)
            if (secDate != null) currentSectionDates = secDate
            i++
            continue
        } else if (line.startsWith("Обороты:", true)) {
            inCancelBlock = false
            inAppointBlock = false
            i++
            continue
        }

        // Ищем объявление поезда (только 4-значные пригородные номера 6000-7999)
        val trainStartRegex = Regex("(?i)^(?:нр\\.?|поезд\\s*№?)\\s*:?\\s*([67]\\d{3}(?:/[67]\\d{3})?)\\b")
        val tm = trainStartRegex.find(line)
        if (tm != null && !inBologoeSection) {
            val trainNum = tm.groupValues[1]

            // Собираем все строки этого поезда до следующего поезда или секции
            val blockLines = mutableListOf(line)
            var j = i + 1
            while (j < lines.size) {
                val nl = lines[j].trim()
                if (nl.isEmpty()) { j++; continue }
                if (trainStartRegex.containsMatchIn(nl) ||
                    Regex("(?i)^(?:нр\\.?|поезд\\s*№?)\\s*\\d{1,4}").containsMatchIn(nl) ||
                    Regex("(?i)(?:^|\\s|[0-9.]+)(?:отменить|не назначать|назначить|изменить(?:\\s+расписание)?|обороты)\\s*[:(]").containsMatchIn(nl) ||
                    nl.startsWith("Четным пригородным", true) ||
                    nl.startsWith("Поздний выход", true)
                ) {
                    break
                }
                blockLines.add(nl)
                j++
            }
            i = j - 1

            val fullBlock = blockLines.joinToString(" ")

            // Проверка на чужие станции (Торжок, Бологое, Угловка)
            if (fullBlock.contains("Торжок", true) || fullBlock.contains("Бологое", true) || fullBlock.contains("Угловка", true)) {
                i++
                continue
            }

            // Извлекаем даты конкретно для этого поезда
            val allParens = Regex("\\(([^)]+)\\)").findAll(line).map { it.groupValues[1] }.toList()
            val dateParen = allParens.firstOrNull { Regex("\\d{1,2}[/.\\-]\\d{1,2}|нояб|окт").containsMatchIn(it.lowercase()) }
            val inlineDate = Regex("(?i)(?:только\\s*\\d{1,2}[/.\\-]\\d{1,2}|с\\s*\\d{1,2}[/.\\-]\\d{1,2}\\s*по\\s*\\d{1,2}[/.\\-]\\d{1,2}[^)]*)").find(line)?.value
            val rawDates = dateParen ?: inlineDate ?: currentSectionDates.ifEmpty { defaultDateGroup }
            val formattedDates = formatTelegramDateGroup(rawDates, defaultDateGroup)

            // Извлекаем частичную отмену ("на участке Клин - Поварово")
            val partialMatch = Regex("(?i)\\(на\\s+участке\\s+([^)]+)\\)").find(line)
            val partialCancel = partialMatch?.groupValues?.get(1)?.trim() ?: ""

            // Извлекаем маршрут из первой строки поезда
            var routePart = line.substring(tm.range.last + 1)
                .replace(":", " ")
                .replace(Regex("\\([^)]*\\)"), " ")
                .replace(Regex("(?i)только\\s*\\d{1,2}[/.\\-]\\d{1,2}.*"), " ")
                .replace(Regex("(?i)с\\s*\\d{1,2}[/.\\-]\\d{1,2}.*"), " ")
                .trim()
            val routeClean = Regex("([А-Яа-яЁё0-9\\s.\\-]+?\\s*[-–—]\\s*[А-Яа-яЁё0-9\\s.\\-]+)").find(routePart)?.groupValues?.get(1)?.trim() ?: ""

            // Извлекаем явно отмененные остановки из примечаний "(Оп Подрезково, Левобережная проследует без остановки)"
            val skippedList = mutableListOf<String>()
            val skipRegex = Regex("(?i)\\(([^)]*?)\\s+проследу(?:ет|ют)\\s+без\\s+остановки\\)")
            for (sm in skipRegex.findAll(fullBlock)) {
                val rawNames = sm.groupValues[1]
                rawNames.split(",").forEach { part ->
                    val norm = normalizeStationKey(part)
                    if (norm.isNotEmpty() && norm !in CLOSED_OR_TRANSIT_STATIONS) {
                        val disp = displayStationName(norm)
                        if (disp !in skippedList) skippedList.add(disp)
                    }
                }
            }

            // Извлекаем явно назначенные остановки "(Оп Поваровка назначается остановка)"
            val appointedList = mutableListOf<String>()
            val appRegex = Regex("(?i)\\(([^)]*?)\\s+назнача(?:ется|ются)\\s+остановк[аи]\\)")
            for (am in appRegex.findAll(fullBlock)) {
                val rawNames = am.groupValues[1]
                rawNames.split(",").forEach { part ->
                    val norm = normalizeStationKey(part)
                    if (norm.isNotEmpty()) {
                        val disp = displayStationName(norm)
                        if (disp !in appointedList) appointedList.add(disp)
                    }
                }
            }

            // Парсим постанционное расписание из тела блока
            val bodyWithoutParens = fullBlock.replace(Regex("\\([^)]*\\)"), " ")
            val stMap = LinkedHashMap<String, Pair<String, Boolean>>()

            val timePointPattern = "\\d{1,2}[:.]\\d{2}(?:[:.]\\d{2})?(?:,\\d+)?"
            val segRegex = Regex("([А-Яа-яЁё][А-Яа-яЁё0-9\\s.\\-]{1,28}?)\\s+(?:приб\\.?\\s*$timePointPattern\\s+подача[^,]+?\\s+)?($timePointPattern)(?:\\s*[-–—:]\\s*($timePointPattern))?")

            val segments = bodyWithoutParens.split(",")
            for ((idx, rawSeg) in segments.withIndex()) {
                val seg = rawSeg.trim()
                if (seg.isEmpty()) continue
                val m = segRegex.find(seg) ?: continue
                val rawSt = m.groupValues[1]
                    .replace(Regex("(?i)действующим\\s+расписанием"), "")
                    .replace(Regex("(?i)до\\s+ст\\.?"), "")
                    .replace(Regex("(?i)далее.*"), "")
                    .trim()
                val normSt = normalizeStationKey(rawSt)
                if (normSt.isEmpty()) continue

                val t1 = m.groupValues[2]
                val t2 = m.groupValues[3]

                // Проверка: была ли в исходном сегменте пометка "(тех. стоянка)" или "(тех)"
                val origSegInFull = fullBlock.substringAfter(rawSt, "").take(45)
                val isTechStop = origSegInFull.contains("тех", true) || seg.contains("без пассажиров", true)

                val depTime = if (t2.isNotEmpty()) cleanSingleTimeToken(t2) else cleanSingleTimeToken(t1)
                val isFirstOrLastSeg = (idx == 0 || idx == segments.lastIndex)
                val hasPassengerStop = (t2.isNotEmpty() || isFirstOrLastSeg || seg.contains("отпр", true) || seg.contains("приб", true)) && !isTechStop

                if (depTime.isNotEmpty()) {
                    stMap[normSt] = Pair(depTime, hasPassengerStop)
                }
            }

            val isFullCancel = inCancelBlock && partialCancel.isEmpty()

            events.add(
                TelegramTrainEvent(
                    trainNum = trainNum,
                    rawDates = rawDates,
                    formattedDateGroup = formattedDates,
                    telegramHeader = telHeader,
                    routeFromTelegram = routeClean,
                    isCancelledFull = isFullCancel,
                    partialCancelSection = partialCancel,
                    isAppointedFull = inAppointBlock,
                    stationTimes = stMap,
                    skippedStops = skippedList,
                    appointedStops = appointedList,
                    trackNotes = "",
                    rawBlockText = fullBlock
                )
            )
        }
        i++
    }
    return events
}

// ==================== ПАРСЕР ФАЙЛОВ «ОТМЕНЫ / ВЫГРУЗКИ» (.TXT, .DOCX) ====================

fun parseOtmenyFiles(files: List<File>): Map<String, String> {
    // Ключ: номерПоезда -> строка статуса из файла "Отмены"
    val map = mutableMapOf<String, String>()
    files.forEach { f ->
        val txt = readDocumentText(f)
        txt.lines().forEach { line ->
            val m = Regex("№\\s*(\\d{4}(?:/\\d{4})?)\\s+(.*?-\\s*.*)").find(line)
            if (m != null) {
                val num = m.groupValues[1].split("/").first()
                map[num] = m.groupValues[2].trim()
            } else {
                val appM = Regex("(?i)Назначается\\s+№\\s*(\\d{4}(?:/\\d{4})?)\\s+(.*)").find(line)
                if (appM != null) {
                    val num = appM.groupValues[1].split("/").first()
                    map[num] = "Назначается ${appM.groupValues[2].trim()}"
                }
            }
        }
    }
    return map
}

// ==================== ЯДРО СРАВНЕНИЯ И АНАЛИТИКИ TRAIN-R ====================

fun buildStationComparison(
    targetStationDisplay: String,
    afishaDb: Map<String, AfishaTrainInfo>,
    telegramEvents: List<TelegramTrainEvent>,
    otmenyMap: Map<String, String>
): Pair<List<StationComparisonRow>, List<StationComparisonRow>> {
    val targetKey = normalizeStationKey(targetStationDisplay)
    val toMoscowRows = mutableListOf<StationComparisonRow>()
    val fromMoscowRows = mutableListOf<StationComparisonRow>()
    val handledTrains = mutableSetOf<String>()

    for (ev in telegramEvents) {
        val baseNum = ev.trainNum.split("/").first()
        val afInfo = afishaDb[baseNum]

        val dir = afInfo?.direction ?: getDirectionByTrainNumber(ev.trainNum)
        val isMcd = afInfo?.isMcd ?: ("/" in ev.trainNum || ev.routeFromTelegram.contains("Ипподром", true) || baseNum.startsWith("72") || baseNum.startsWith("73") || baseNum.startsWith("74"))

        val routeStr = when {
            afInfo != null -> afInfo.routeString
            ev.routeFromTelegram.isNotEmpty() -> ev.routeFromTelegram
            dir == Direction.TO_MOSCOW -> if (isMcd) "Крюково – Ипподром" else "Крюково – Москва"
            else -> if (isMcd) "Ипподром – Крюково" else "Москва – Крюково"
        }

        val destLabel = afInfo?.destinationLabel ?: if (dir == Direction.TO_MOSCOW) {
            if (isMcd) "до Ипподрома (47 км)" else "до Москвы"
        } else {
            when {
                routeStr.contains("Твер", true) -> "до Твери"
                routeStr.contains("Клин", true) -> "до Клина"
                routeStr.contains("Конаков", true) -> "до Конаково ГРЭС"
                routeStr.contains("Подсолнеч", true) -> "до Подсолнечной"
                routeStr.contains("Решетник", true) -> "до Решетниково"
                else -> "до Крюково"
            }
        }

        val afTimeAtStation = afInfo?.stationTimes?.get(targetKey) ?: ""
        val telPairAtStation = ev.stationTimes[targetKey]
        val telTimeAtStation = telPairAtStation?.first ?: ""
        val telHasStop = telPairAtStation?.second ?: false

        val isStationInSkippedList = ev.skippedStops.any { normalizeStationKey(it) == targetKey }
        val isStationInAppointedList = ev.appointedStops.any { normalizeStationKey(it) == targetKey }

        // Проверяем, проходит ли этот поезд вообще через выбранную станцию
        val passesStation = afTimeAtStation.isNotEmpty() || telPairAtStation != null || isStationInSkippedList || isStationInAppointedList ||
                (afInfo == null && (ev.isCancelledFull || ev.isAppointedFull))
        if (!passesStation) continue

        var statusText = ""
        var isCanc = false
        var isPassNoStop = false
        var isApp = false

        // 1. Полная отмена поезда
        if (ev.isCancelledFull) {
            isCanc = true
            statusText = "ОТМЕНЕН"
        }
        // 2. Частичная отмена на участке ("на участке Клин - Решетниково")
        else if (ev.partialCancelSection.isNotEmpty()) {
            val secNorm = ev.partialCancelSection.lowercase()
            if ("крюково" in secNorm && "подсолнеч" in secNorm && dir == Direction.FROM_MOSCOW) {
                statusText = if (telTimeAtStation.isNotEmpty() && telTimeAtStation != afTimeAtStation) {
                    "отправится в $telTimeAtStation следует до Крюково"
                } else {
                    "следует до Крюково"
                }
            } else if ("клин" in secNorm && "решетник" in secNorm && dir == Direction.FROM_MOSCOW) {
                statusText = "следует до Клина"
            } else if ("поваров" in secNorm && "клин" in secNorm && dir == Direction.FROM_MOSCOW) {
                statusText = "следует до Поварово-1"
            } else {
                isCanc = true
                statusText = "ОТМЕНЕН (на уч. ${ev.partialCancelSection})"
            }
        }
        // 3. Назначение нового поезда
        else if (ev.isAppointedFull) {
            if (telTimeAtStation.isEmpty() && afTimeAtStation.isEmpty()) continue
            isApp = true
            val t = telTimeAtStation.ifEmpty { afTimeAtStation }
            statusText = "НАЗНАЧЕН в $t"
        }
        // 4. Изменение расписания / остановок
        else {
            if (isStationInAppointedList) {
                isApp = true
                val t = telTimeAtStation.ifEmpty { afTimeAtStation }
                statusText = "Назначена остановка $targetStationDisplay, отправится в $t"
            } else if (isStationInSkippedList || (telPairAtStation != null && !telHasStop && afTimeAtStation.isNotEmpty())) {
                isCanc = true
                isPassNoStop = true
                statusText = "ОТМЕНЕН (проследует без остановки)"
            } else if (telTimeAtStation.isNotEmpty() && telHasStop) {
                if (afTimeAtStation.isNotEmpty() && telTimeAtStation != afTimeAtStation) {
                    statusText = "отправится в $telTimeAtStation"
                } else if (afTimeAtStation.isEmpty()) {
                    statusText = "отправится в $telTimeAtStation"
                }
            }

            // Если сам поезд на данной станции не отменен, но дальше по маршруту у него отменены или назначены остановки — дописываем как в эталоне Сходни!
            if (!isCanc) {
                val otherSkipped = ev.skippedStops.filter { normalizeStationKey(it) != targetKey }
                val otherAppointed = ev.appointedStops.filter { normalizeStationKey(it) != targetKey }

                if (otherSkipped.isNotEmpty()) {
                    val suffix = "Отменены остановки: ${otherSkipped.joinToString(", ")}."
                    statusText = if (statusText.isNotEmpty()) "$statusText. $suffix" else suffix
                }
                if (otherAppointed.isNotEmpty()) {
                    val suffix = "Назначены остановки: ${otherAppointed.joinToString(", ")}."
                    statusText = if (statusText.isNotEmpty()) "$statusText. $suffix" else suffix
                }
            }
        }

        val otmStatus = otmenyMap[baseNum] ?: ""

        // Если ни по телеграмме, ни по файлу "Отмены" изменений для этой станции нет — пропускаем
        if (statusText.isEmpty() && otmStatus.isEmpty()) continue

        val hasDiscrep = otmStatus.isNotEmpty() && statusText.isNotEmpty() &&
                !otmStatus.contains(telTimeAtStation.ifEmpty { "___" }) &&
                !(otmStatus.contains("ОТМЕНЕН", true) && isCanc)

        val row = StationComparisonRow(
            dateGroup = ev.formattedDateGroup,
            trainNum = ev.trainNum,
            daysFlag = afInfo?.daysFlag ?: "",
            routeString = routeStr,
            destinationLabel = destLabel,
            isMcd = isMcd,
            direction = dir,
            stationName = targetStationDisplay,
            telegramHeader = ev.telegramHeader,
            afishaTime = afTimeAtStation.ifEmpty { "-" },
            telegramTime = telTimeAtStation.ifEmpty { if (isCanc) "ОТМ" else "-" },
            otmenyFileStatus = otmStatus,
            finalChangeText = statusText.ifEmpty { otmStatus },
            isCancelledAtStation = isCanc,
            isPassWithoutStop = isPassNoStop,
            isAppointedAtStation = isApp,
            skippedStopsSummary = ev.skippedStops.joinToString(", "),
            appointedStopsSummary = ev.appointedStops.joinToString(", "),
            hasDiscrepancyWithOtmeny = hasDiscrep,
            fullAfishaStations = afInfo?.stationTimes ?: emptyMap(),
            fullTelegramStations = ev.stationTimes,
            rawTelegramText = ev.rawBlockText
        )

        if (dir == Direction.TO_MOSCOW) toMoscowRows.add(row) else fromMoscowRows.add(row)
        handledTrains.add(baseNum)
    }

    // Добавляем поезда, которые есть в файле "Отмены", но отсутствуют в телеграмме (обратная сверка)
    for ((num, otmStat) in otmenyMap) {
        if (num in handledTrains) continue
        val afInfo = afishaDb[num]
        val afTime = afInfo?.stationTimes?.get(targetKey) ?: "-"
        val dir = afInfo?.direction ?: getDirectionByTrainNumber(num)
        val isMcd = afInfo?.isMcd ?: (num.startsWith("72") || num.startsWith("73") || num.startsWith("74"))
        val routeStr = afInfo?.routeString ?: if (dir == Direction.TO_MOSCOW) "Крюково – Москва" else "Москва – Крюково"
        val destLabel = afInfo?.destinationLabel ?: if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Крюково"

        val row = StationComparisonRow(
            dateGroup = "По файлу «Отмены»",
            trainNum = num,
            daysFlag = afInfo?.daysFlag ?: "",
            routeString = routeStr,
            destinationLabel = destLabel,
            isMcd = isMcd,
            direction = dir,
            stationName = targetStationDisplay,
            telegramHeader = "Нет в телеграмме!",
            afishaTime = afTime,
            telegramTime = "-",
            otmenyFileStatus = otmStat,
            finalChangeText = otmStat,
            isCancelledAtStation = otmStat.contains("ОТМЕНЕН", true),
            isPassWithoutStop = otmStat.contains("без остановки", true),
            isAppointedAtStation = otmStat.contains("Назнач", true),
            skippedStopsSummary = "",
            appointedStopsSummary = "",
            hasDiscrepancyWithOtmeny = true,
            fullAfishaStations = afInfo?.stationTimes ?: emptyMap(),
            fullTelegramStations = emptyMap(),
            rawTelegramText = "Поезд присутствует в файле выгрузки «Отмены», но не найден в загруженных телеграммах."
        )
        if (dir == Direction.TO_MOSCOW) toMoscowRows.add(row) else fromMoscowRows.add(row)
    }

    val comparator = compareBy<StationComparisonRow>(
        { extractFirstDayForSort(it.dateGroup) },
        { if (it.afishaTime != "-") it.afishaTime else it.telegramTime }
    )

    return Pair(
        toMoscowRows.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.finalChangeText}" }.sortedWith(comparator),
        fromMoscowRows.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.finalChangeText}" }.sortedWith(comparator)
    )
}

// ==================== ЭКСПОРТ ПО ЭТАЛОНУ «СХОДНЯ НА МОСКВУ / ОТ МОСКВЫ» ====================

fun exportStationFilesToFolders(
    station: String,
    toMoscow: List<StationComparisonRow>,
    fromMoscow: List<StationComparisonRow>
) {
    val desktop = File(System.getProperty("user.home"), "Desktop")
    val dirToMoscow = File(desktop, "До Москвы и Ипподрома (нечетные)").apply { if (!exists()) mkdirs() }
    val dirFromMoscow = File(desktop, "От Москвы и Ипподрома (четные)").apply { if (!exists()) mkdirs() }

    if (toMoscow.isNotEmpty()) {
        val f = File(dirToMoscow, "$station на Москву.txt")
        f.printWriter(Charsets.UTF_8).use { out ->
            out.println("$station    Изменения в графике движения эл/поездов")
            out.println("на Ленинградский вокзал, Ипподром")
            val grouped = toMoscow.groupBy { it.dateGroup }
            for (dKey in grouped.keys.sortedBy { extractFirstDayForSort(it) }) {
                out.println(dKey)
                for (r in grouped[dKey] ?: emptyList()) {
                    if (r.isAppointedAtStation && r.afishaTime == "-") {
                        out.println("        Назначается №${r.trainNum} ${r.destinationLabel} в ${r.telegramTime}")
                    } else {
                        val depTime = if (r.afishaTime != "-") r.afishaTime else r.telegramTime
                        out.println("        №${r.trainNum} ${r.destinationLabel} отправлением в $depTime - ${r.finalChangeText}")
                    }
                }
            }
        }
    }

    if (fromMoscow.isNotEmpty()) {
        val f = File(dirFromMoscow, "$station от Москвы.txt")
        f.printWriter(Charsets.UTF_8).use { out ->
            out.println("$station   Изменения в графике движения эл/поездов")
            out.println("от Ленинградского вокзала, от Ипподрома")
            val grouped = fromMoscow.groupBy { it.dateGroup }
            for (dKey in grouped.keys.sortedBy { extractFirstDayForSort(it) }) {
                out.println(dKey)
                for (r in grouped[dKey] ?: emptyList()) {
                    if (r.isAppointedAtStation && r.afishaTime == "-") {
                        out.println("        Назначается №${r.trainNum} ${r.destinationLabel} в ${r.telegramTime}")
                    } else {
                        val depTime = if (r.afishaTime != "-") r.afishaTime else r.telegramTime
                        out.println("        №${r.trainNum} ${r.destinationLabel} отправлением в $depTime - ${r.finalChangeText}")
                    }
                }
            }
        }
    }
}

// ==================== ВЕКТОРНЫЙ ЛОГОТИП В СТАРИННОМ СТИЛЕ «TRAIN-R» ====================

@Composable
fun VintageTrainLogo(modifier: Modifier = Modifier.size(46.dp)) {
    Canvas(modifier = modifier) {
        val w = size.width
        val h = size.height
        // Темно-синий медальон с золотым ободом
        drawCircle(color = Color(0xFF0F172A), radius = w / 2f)
        drawCircle(color = Color(0xFFD97706), radius = w / 2f - 2f, style = Stroke(width = 3f))

        // Клубы пара сверху
        drawCircle(color = Color(0xFFE2E8F0), radius = w * 0.08f, center = Offset(w * 0.42f, h * 0.25f))
        drawCircle(color = Color(0xFFF8FAFC), radius = w * 0.10f, center = Offset(w * 0.55f, h * 0.22f))
        drawCircle(color = Color(0xFFE2E8F0), radius = w * 0.07f, center = Offset(w * 0.66f, h * 0.28f))

        // Кабина и котел паровоза
        drawRect(color = Color(0xFFB91C1C), topLeft = Offset(w * 0.22f, h * 0.42f), size = Size(w * 0.20f, h * 0.26f))
        drawRect(color = Color(0xFFFEF3C7), topLeft = Offset(w * 0.26f, h * 0.47f), size = Size(w * 0.11f, h * 0.11f))
        drawRect(color = Color(0xFF1E293B), topLeft = Offset(w * 0.42f, h * 0.48f), size = Size(w * 0.30f, h * 0.20f))

        // Дымовая труба
        drawRect(color = Color(0xFFD97706), topLeft = Offset(w * 0.60f, h * 0.34f), size = Size(w * 0.08f, h * 0.14f))

        // Скотоотбрасыватель спереди
        val cowPath = Path().apply {
            moveTo(w * 0.72f, h * 0.56f)
            lineTo(w * 0.82f, h * 0.68f)
            lineTo(w * 0.72f, h * 0.68f)
            close()
        }
        drawPath(cowPath, color = Color(0xFFD97706))

        // Колеса
        drawCircle(color = Color(0xFFD97706), radius = w * 0.08f, center = Offset(w * 0.32f, h * 0.72f))
        drawCircle(color = Color(0xFFD97706), radius = w * 0.08f, center = Offset(w * 0.50f, h * 0.72f))
        drawCircle(color = Color(0xFFD97706), radius = w * 0.06f, center = Offset(w * 0.66f, h * 0.74f))

        // Рельса
        drawLine(color = Color(0xFFF59E0B), start = Offset(w * 0.16f, h * 0.81f), end = Offset(w * 0.84f, h * 0.81f), strokeWidth = 2.5f)
    }
}

// ==================== ГЛАВНЫЙ ИНТЕРФЕЙС TRAIN-R ====================

@Composable
fun TrainRApp() {
    val cachedAfisha = remember { getCachedAfishaFile() }
    var afishaFile by remember { mutableStateOf<File?>(if (cachedAfisha.exists()) cachedAfisha else null) }
    var telegramFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var otmenyFiles by remember { mutableStateOf<List<File>>(emptyList()) }

    var afishaDb by remember { mutableStateOf<Map<String, AfishaTrainInfo>>(emptyMap()) }
    var telegramEvents by remember { mutableStateOf<List<TelegramTrainEvent>>(emptyList()) }
    var otmenyMap by remember { mutableStateOf<Map<String, String>>(emptyMap()) }

    var currentStation by remember { mutableStateOf("Сходня") }
    var selectedStationsForExport by remember { mutableStateOf(setOf("Сходня", "Химки", "Зеленоград-Крюково", "Подрезково", "Моссельмаш")) }

    var toMoscowList by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }
    var fromMoscowList by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }

    var activeTab by remember { mutableStateOf(0) } // 0: На Москву, 1: От Москвы, 2: Отчет по станции, 3: Глобальный поиск
    var searchQuery by remember { mutableStateOf("") }
    var globalSearchResults by remember { mutableStateOf<List<StationComparisonRow>>(emptyList()) }
    var inspectedTrainRow by remember { mutableStateOf<StationComparisonRow?>(null) }

    var isBusy by remember { mutableStateOf(false) }
    var statusText by remember { mutableStateOf("Добро пожаловать в Train-R! Загрузите Афишу (.xlsx), Телеграммы (.docx/.pdf) или файл «Отмены».") }

    val scope = rememberCoroutineScope()

    // Автозагрузка кэшированной Афиши при старте
    LaunchedEffect(afishaFile) {
        if (afishaFile != null && afishaFile!!.exists()) {
            withContext(Dispatchers.IO) {
                val loaded = loadAfishaDatabase(afishaFile)
                afishaDb = loaded
                statusText = "Базовая Афиша активна (распознано ${loaded.size} ниток поездов с маршрутами)."
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

    fun recalculateStation(stName: String) {
        scope.launch {
            isBusy = true
            statusText = "Анализ расписания и сверка изменений по станции $stName..."
            val (toM, fromM) = withContext(Dispatchers.IO) {
                val parsedTels = telegramFiles.flatMap { parseTelegramDocument(it) }
                val parsedOtm = parseOtmenyFiles(otmenyFiles)
                telegramEvents = parsedTels
                otmenyMap = parsedOtm
                buildStationComparison(stName, afishaDb, parsedTels, parsedOtm)
            }
            toMoscowList = toM
            fromMoscowList = fromM
            isBusy = false
            statusText = "Станция $stName: найдено изменений На Москву/Ипподром — ${toM.size}, От Москвы — ${fromM.size}."
        }
    }

    fun runGlobalTrainSearch(query: String) {
        val q = query.trim()
        if (q.isEmpty()) return
        scope.launch {
            isBusy = true
            statusText = "Глобальный поиск поезда №$q по всем файлам (Афиша, Телеграммы, Выгрузки)..."
            val found = withContext(Dispatchers.IO) {
                val parsedTels = if (telegramEvents.isEmpty() && telegramFiles.isNotEmpty()) telegramFiles.flatMap { parseTelegramDocument(it) } else telegramEvents
                val parsedOtm = if (otmenyMap.isEmpty() && otmenyFiles.isNotEmpty()) parseOtmenyFiles(otmenyFiles) else otmenyMap
                val results = mutableListOf<StationComparisonRow>()

                for (st in POLYGON_STATIONS) {
                    val (t, f) = buildStationComparison(st, afishaDb, parsedTels, parsedOtm)
                    results.addAll((t + f).filter { it.trainNum.contains(q, true) })
                }

                // Если поезда нет в изменениях телеграмм, но он есть в базовой Афише — покажем его полное досье из Афиши!
                if (results.isEmpty()) {
                    afishaDb.values.filter { it.trainNum.contains(q, true) }.forEach { af ->
                        results.add(
                            StationComparisonRow(
                                dateGroup = "Действующий график (Афиша)",
                                trainNum = af.trainNum,
                                daysFlag = af.daysFlag,
                                routeString = af.routeString,
                                destinationLabel = af.destinationLabel,
                                isMcd = af.isMcd,
                                direction = af.direction,
                                stationName = "Все станции маршрута",
                                telegramHeader = "Без изменений в телеграммах",
                                afishaTime = af.stationTimes.values.firstOrNull() ?: "-",
                                telegramTime = "-",
                                otmenyFileStatus = "",
                                finalChangeText = "Следует действующим расписанием Афиши (${af.stationTimes.size} ост.)",
                                isCancelledAtStation = false,
                                isPassWithoutStop = false,
                                isAppointedAtStation = false,
                                skippedStopsSummary = "",
                                appointedStopsSummary = "",
                                hasDiscrepancyWithOtmeny = false,
                                fullAfishaStations = af.stationTimes,
                                fullTelegramStations = emptyMap(),
                                rawTelegramText = "Поезд №${af.trainNum} (${af.routeString}) найден в базовой Афише. В загруженных телеграммах изменений по нему не объявлено."
                            )
                        )
                    }
                }
                results
            }
            globalSearchResults = found
            activeTab = 3
            isBusy = false
            statusText = "Поиск №$q завершен: найдено записей — ${found.size}. Нажмите на строку для открытия полного досье поезда."
        }
    }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = TrainRTheme.BgMain) {
            Column(modifier = Modifier.fillMaxSize().padding(10.dp)) {

                // 1. КОМПАКТНАЯ ШАПКА С ВИНТАЖНЫМ ЛОГОТИПОМ TRAIN-R И ФАЙЛАМИ
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(horizontal = 12.dp, vertical = 8.dp),
                        verticalAlignment = Alignment.CenterVertically,
                        horizontalArrangement = Arrangement.spacedBy(10.dp)
                    ) {
                        VintageTrainLogo()
                        Column(modifier = Modifier.width(260.dp)) {
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Text("Train-R", fontSize = 20.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                                Spacer(modifier = Modifier.width(6.dp))
                                Surface(shape = RoundedCornerShape(4.dp), color = Color(0xFFFEF3C7)) {
                                    Text("МТППК • МЦД-3", fontSize = 9.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.GoldAccent, modifier = Modifier.padding(horizontal = 5.dp, vertical = 2.dp))
                                }
                            }
                            Text(statusText, fontSize = 10.sp, fontWeight = FontWeight.SemiBold, color = TrainRTheme.TextSub, maxLines = 2, overflow = TextOverflow.Ellipsis)
                        }

                        // Компактные слоты загрузки 3 типов файлов с кнопками очистки
                        CompactFileBadge(
                            title = "1. АФИША (.XLSX)",
                            subtitle = if (afishaFile != null) "Ниток в базе: ${afishaDb.size}" else "Не загружена",
                            isLoaded = afishaFile != null,
                            accent = TrainRTheme.GreenSuccess,
                            modifier = Modifier.weight(1f),
                            onSelect = {
                                val f = openFileDialog("Выберите файл Афиши (.xlsx)", false)
                                if (f.isNotEmpty()) {
                                    try {
                                        FileInputStream(f.first()).use { inp -> FileOutputStream(cachedAfisha).use { out -> inp.copyTo(out) } }
                                        afishaFile = cachedAfisha
                                    } catch (_: Exception) { afishaFile = f.first() }
                                }
                            },
                            onClear = {
                                if (cachedAfisha.exists()) cachedAfisha.delete()
                                afishaFile = null
                                afishaDb = emptyMap()
                                statusText = "Афиша удалена из памяти."
                            }
                        )

                        CompactFileBadge(
                            title = "2. ТЕЛЕГРАММЫ (${telegramFiles.size})",
                            subtitle = if (telegramFiles.isNotEmpty()) telegramFiles.joinToString { it.nameWithoutExtension.take(10) } else "Word / PDF (можно неск.)",
                            isLoaded = telegramFiles.isNotEmpty(),
                            accent = TrainRTheme.RoyalBlue,
                            modifier = Modifier.weight(1f),
                            onSelect = {
                                val f = openFileDialog("Выберите одну или несколько телеграмм", true)
                                if (f.isNotEmpty()) {
                                    telegramFiles = telegramFiles + f
                                    recalculateStation(currentStation)
                                }
                            },
                            onClear = {
                                telegramFiles = emptyList()
                                telegramEvents = emptyList()
                                statusText = "Список телеграмм очищен."
                            }
                        )

                        CompactFileBadge(
                            title = "3. ФАЙЛ «ОТМЕНЫ» (${otmenyFiles.size})",
                            subtitle = if (otmenyFiles.isNotEmpty()) "Загружено: ${otmenyFiles.size} док." else "TXT / DOCX для сверки",
                            isLoaded = otmenyFiles.isNotEmpty(),
                            accent = Color(0xFF7C3AED),
                            modifier = Modifier.weight(1f),
                            onSelect = {
                                val f = openFileDialog("Выберите файлы выгрузки «Отмены» (.txt, .docx)", true)
                                if (f.isNotEmpty()) {
                                    otmenyFiles = f
                                    recalculateStation(currentStation)
                                }
                            },
                            onClear = {
                                otmenyFiles = emptyList()
                                otmenyMap = emptyMap()
                                statusText = "Файлы «Отмены» очищены."
                            }
                        )
                    }
                }

                Spacer(modifier = Modifier.height(6.dp))

                // 2. ПАНЕЛЬ ВЫБОРА СТАНЦИИ И ГЕНЕРАЦИИ ФАЙЛОВ В 2 ПАПКИ
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg),
                    elevation = CardDefaults.cardElevation(1.dp)
                ) {
                    Column(modifier = Modifier.fillMaxWidth().padding(8.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                            Text("Станция:", fontWeight = FontWeight.ExtraBold, fontSize = 12.sp, color = TrainRTheme.NavyDark)

                            FilledTonalButton(
                                onClick = {
                                    selectedStationsForExport = if (selectedStationsForExport.size == POLYGON_STATIONS.size) setOf(currentStation) else POLYGON_STATIONS.toSet()
                                },
                                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 2.dp),
                                modifier = Modifier.height(30.dp),
                                shape = RoundedCornerShape(6.dp)
                            ) {
                                Text(if (selectedStationsForExport.size == POLYGON_STATIONS.size) "Снять галочки" else "Выбрать все (${POLYGON_STATIONS.size})", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }

                            Spacer(modifier = Modifier.weight(1f))

                            Button(
                                onClick = { recalculateStation(currentStation) },
                                enabled = !isBusy,
                                colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RoyalBlue),
                                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 4.dp),
                                modifier = Modifier.height(32.dp),
                                shape = RoundedCornerShape(6.dp)
                            ) {
                                Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(15.dp))
                                Spacer(modifier = Modifier.width(4.dp))
                                Text("Сверить и показать: $currentStation", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }

                            Button(
                                onClick = {
                                    scope.launch {
                                        isBusy = true
                                        statusText = "Генерация выгрузок в 2 папки на Рабочем столе..."
                                        withContext(Dispatchers.IO) {
                                            val pTels = if (telegramEvents.isEmpty()) telegramFiles.flatMap { parseTelegramDocument(it) } else telegramEvents
                                            val pOtm = if (otmenyMap.isEmpty()) parseOtmenyFiles(otmenyFiles) else otmenyMap
                                            selectedStationsForExport.forEach { st ->
                                                val (tM, fM) = buildStationComparison(st, afishaDb, pTels, pOtm)
                                                exportStationFilesToFolders(st, tM, fM)
                                            }
                                        }
                                        isBusy = false
                                        statusText = "ГОТОВО! Выгрузки (${selectedStationsForExport.size} ст.) сохранены на Рабочем столе в 2 папках («До Москвы...» и «От Москвы...»)."
                                    }
                                },
                                enabled = !isBusy && selectedStationsForExport.isNotEmpty(),
                                colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.RedCancel),
                                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 4.dp),
                                modifier = Modifier.height(32.dp),
                                shape = RoundedCornerShape(6.dp)
                            ) {
                                Icon(Icons.Default.Share, contentDescription = null, modifier = Modifier.size(15.dp))
                                Spacer(modifier = Modifier.width(4.dp))
                                Text("Генерировать выгрузки в 2 папки (${selectedStationsForExport.size} ст.)", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }
                        }

                        Spacer(modifier = Modifier.height(6.dp))

                        // Компактная двухрядная сетка станций
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                            POLYGON_STATIONS.take(11).forEach { st ->
                                StationChipItem(
                                    name = st,
                                    isChecked = st in selectedStationsForExport,
                                    isSelected = currentStation == st,
                                    modifier = Modifier.weight(1f),
                                    onCheck = { selectedStationsForExport = if (st in selectedStationsForExport) selectedStationsForExport - st else selectedStationsForExport + st },
                                    onSelect = { currentStation = st; recalculateStation(st) }
                                )
                            }
                        }
                        Spacer(modifier = Modifier.height(3.dp))
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                            POLYGON_STATIONS.drop(11).forEach { st ->
                                StationChipItem(
                                    name = st,
                                    isChecked = st in selectedStationsForExport,
                                    isSelected = currentStation == st,
                                    modifier = Modifier.weight(1f),
                                    onCheck = { selectedStationsForExport = if (st in selectedStationsForExport) selectedStationsForExport - st else selectedStationsForExport + st },
                                    onSelect = { currentStation = st; recalculateStation(st) }
                                )
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(6.dp))

                // 3. ПАНЕЛЬ ВКЛАДОК И ПОИСКА ПО НОМЕРУ ПОЕЗДА
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    TabRow(
                        selectedTabIndex = activeTab,
                        containerColor = TrainRTheme.CardBg,
                        contentColor = TrainRTheme.RoyalBlue,
                        modifier = Modifier.width(680.dp).height(42.dp).clip(RoundedCornerShape(8.dp))
                    ) {
                        Tab(selected = activeTab == 0, onClick = { activeTab = 0 }, text = { Text("ДО МОСКВЫ / ИППОДРОМА (${toMoscowList.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp) })
                        Tab(selected = activeTab == 1, onClick = { activeTab = 1 }, text = { Text("ОТ МОСКВЫ НА ОБЛАСТЬ (${fromMoscowList.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp) })
                        Tab(selected = activeTab == 2, onClick = { activeTab = 2 }, text = { Text("ОТЧЕТНОСТЬ ПО СТАНЦИИ", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.GreenSuccess) })
                        Tab(selected = activeTab == 3, onClick = { activeTab = 3 }, text = { Text("ПОИСК ПОЕЗДА (${globalSearchResults.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.GoldAccent) })
                    }

                    Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                        OutlinedTextField(
                            value = searchQuery,
                            onValueChange = { searchQuery = it },
                            placeholder = { Text("№ поезда (напр. 6712, 7529)...", fontSize = 12.sp) },
                            modifier = Modifier.width(240.dp).height(48.dp),
                            singleLine = true,
                            textStyle = LocalTextStyle.current.copy(fontSize = 13.sp, fontWeight = FontWeight.Bold),
                            shape = RoundedCornerShape(8.dp)
                        )
                        Button(
                            onClick = { runGlobalTrainSearch(searchQuery) },
                            colors = ButtonDefaults.buttonColors(containerColor = TrainRTheme.NavyDark),
                            modifier = Modifier.height(44.dp),
                            shape = RoundedCornerShape(8.dp)
                        ) {
                            Icon(Icons.Default.Search, contentDescription = null, modifier = Modifier.size(16.dp))
                            Spacer(modifier = Modifier.width(4.dp))
                            Text("Найти везде", fontSize = 12.sp, fontWeight = FontWeight.Bold)
                        }
                    }
                }

                Spacer(modifier = Modifier.height(6.dp))

                // 4. ГЛАВНАЯ ОБЛАСТЬ ТАБЛИЦЫ ПОЕЗДОВ И ОТЧЕТНОСТИ (МАКСИМУМ МЕСТА)
                Card(
                    modifier = Modifier.fillMaxSize(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = TrainRTheme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    when (activeTab) {
                        0, 1 -> {
                            val baseList = if (activeTab == 0) toMoscowList else fromMoscowList
                            val filtered = if (searchQuery.isBlank()) baseList else baseList.filter { it.trainNum.contains(searchQuery.trim(), true) }
                            TrainTableSection(
                                rows = filtered,
                                emptyMessage = "Нет изменений по ст. $currentStation. Загрузите файлы и нажмите «Сверить и показать: $currentStation».",
                                onRowClick = { inspectedTrainRow = it }
                            )
                        }
                        2 -> {
                            StationAnalyticsReportView(
                                stationName = currentStation,
                                toMoscow = toMoscowList,
                                fromMoscow = fromMoscowList,
                                onTrainClick = { inspectedTrainRow = it }
                            )
                        }
                        3 -> {
                            TrainTableSection(
                                rows = globalSearchResults,
                                emptyMessage = "Введите номер поезда справа вверху и нажмите «Найти везде».",
                                showStationColumn = true,
                                onRowClick = { inspectedTrainRow = it }
                            )
                        }
                    }
                }
            }

            // МОДАЛЬНОЕ ОКНО: ПОЛНОЕ ДОСЬЕ КОНКРЕТНОГО ПОЕЗДА
            if (inspectedTrainRow != null) {
                TrainDossierDialog(
                    row = inspectedTrainRow!!,
                    onDismiss = { inspectedTrainRow = null }
                )
            }
        }
    }
}

// ==================== ТАБЛИЦА ПОЕЗДОВ (СВЕРКА АФИША / ТЕЛЕГРАММА / ОТМЕНЫ) ====================

@Composable
fun TrainTableSection(
    rows: List<StationComparisonRow>,
    emptyMessage: String,
    showStationColumn: Boolean = false,
    onRowClick: (StationComparisonRow) -> Unit
) {
    Column(modifier = Modifier.fillMaxSize().padding(8.dp)) {
        Row(
            modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFE2E8F0)).padding(vertical = 8.dp, horizontal = 8.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            Text("ДАТЫ (ПЕРИОД)", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(135.dp))
            if (showStationColumn) {
                Text("СТАНЦИЯ", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(105.dp))
            }
            Text("№ ПОЕЗДА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(105.dp))
            Text("МАРШРУТ (ОТ – ДО)", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(155.dp))
            Text("ТЕЛЕГРАММА", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(125.dp))
            Text("АФИША", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(65.dp))
            Text("ТЕЛЕГР.", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(65.dp))
            Text("ФАКТИЧЕСКИЕ ИЗМЕНЕНИЯ (ОСТАНОВКИ, ВРЕМЯ, ПУТИ)", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.weight(1f))
            Text("СВЕРКА «ОТМЕНЫ»", fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(150.dp))
        }

        Spacer(modifier = Modifier.height(4.dp))

        if (rows.isEmpty()) {
            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                Text(emptyMessage, fontSize = 13.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.TextSub)
            }
        } else {
            LazyColumn(modifier = Modifier.fillMaxSize()) {
                items(rows) { item ->
                    val bg = when {
                        item.hasDiscrepancyWithOtmeny -> Color(0xFFFEF3C7)
                        item.isCancelledAtStation -> Color(0xFFFFF1F2)
                        item.isAppointedAtStation -> Color(0xFFECFDF5)
                        else -> Color.White
                    }
                    val borderCol = when {
                        item.hasDiscrepancyWithOtmeny -> Color(0xFFF59E0B)
                        item.isCancelledAtStation -> Color(0xFFFECDD3)
                        item.isAppointedAtStation -> Color(0xFFA7F3D0)
                        else -> Color(0xFFE2E8F0)
                    }

                    Row(
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(vertical = 2.dp)
                            .clip(RoundedCornerShape(6.dp))
                            .background(bg)
                            .border(1.dp, borderCol, RoundedCornerShape(6.dp))
                            .clickable { onRowClick(item) }
                            .padding(vertical = 6.dp, horizontal = 8.dp),
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Text(item.dateGroup, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.RoyalBlue, modifier = Modifier.width(135.dp))

                        if (showStationColumn) {
                            Text(item.stationName, fontWeight = FontWeight.Black, fontSize = 11.sp, color = TrainRTheme.NavyDark, modifier = Modifier.width(105.dp))
                        }

                        // Номер поезда + плашка МЦД
                        Row(modifier = Modifier.width(105.dp), verticalAlignment = Alignment.CenterVertically) {
                            Text("№${item.trainNum}", fontWeight = FontWeight.Black, fontFamily = FontFamily.Monospace, fontSize = 13.sp, color = TrainRTheme.NavyDark)
                            if (item.isMcd) {
                                Spacer(modifier = Modifier.width(4.dp))
                                Surface(shape = RoundedCornerShape(3.dp), color = Color(0xFFFFE4E6)) {
                                    Text("МЦД", fontSize = 8.sp, fontWeight = FontWeight.Black, color = Color(0xFFBE123C), modifier = Modifier.padding(horizontal = 3.dp, vertical = 1.dp))
                                }
                            }
                        }

                        // Полный маршрут (От - До)
                        Column(modifier = Modifier.width(155.dp)) {
                            Text(item.routeString, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.NavyDark, maxLines = 1, overflow = TextOverflow.Ellipsis)
                            Text(item.destinationLabel, fontWeight = FontWeight.SemiBold, fontSize = 9.sp, color = TrainRTheme.RoyalBlue)
                        }

                        // Номер и дата телеграммы
                        Text(item.telegramHeader, fontWeight = FontWeight.Bold, fontSize = 10.sp, color = TrainRTheme.TextSub, modifier = Modifier.width(125.dp), maxLines = 2)

                        // Время по Афише (Было)
                        Text(item.afishaTime, fontFamily = FontFamily.Monospace, fontWeight = FontWeight.ExtraBold, fontSize = 12.sp, color = TrainRTheme.TextSub, modifier = Modifier.width(65.dp))

                        // Время по Телеграмме (Стало)
                        Text(
                            text = item.telegramTime,
                            fontFamily = FontFamily.Monospace,
                            fontWeight = FontWeight.Black,
                            fontSize = 12.sp,
                            color = if (item.isCancelledAtStation) TrainRTheme.RedCancel else TrainRTheme.RoyalBlue,
                            modifier = Modifier.width(65.dp)
                        )

                        // Что изменено согласно телеграмме
                        Text(
                            text = item.finalChangeText,
                            fontWeight = FontWeight.ExtraBold,
                            fontSize = 11.sp,
                            color = when {
                                item.isCancelledAtStation -> TrainRTheme.RedCancel
                                item.isAppointedAtStation -> TrainRTheme.GreenSuccess
                                else -> TrainRTheme.TextDark
                            },
                            modifier = Modifier.weight(1f).padding(end = 6.dp)
                        )

                        // Статус из файла «Отмены»
                        Text(
                            text = item.otmenyFileStatus.ifEmpty { "—" },
                            fontWeight = FontWeight.Bold,
                            fontSize = 10.sp,
                            color = if (item.hasDiscrepancyWithOtmeny) TrainRTheme.RedCancel else TrainRTheme.TextSub,
                            modifier = Modifier.width(150.dp),
                            maxLines = 2
                        )
                    }
                }
            }
        }
    }
}

// ==================== ВКЛАДКА ОТЧЕТНОСТИ ПО СТАНЦИИ ====================

@Composable
fun StationAnalyticsReportView(
    stationName: String,
    toMoscow: List<StationComparisonRow>,
    fromMoscow: List<StationComparisonRow>,
    onTrainClick: (StationComparisonRow) -> Unit
) {
    val all = (toMoscow + fromMoscow).distinctBy { "${it.trainNum}_${it.direction}" }
    val cancelled = all.filter { it.isCancelledAtStation }
    val appointed = all.filter { it.isAppointedAtStation }
    val shifted = all.filter { !it.isCancelledAtStation && !it.isAppointedAtStation && it.finalChangeText.contains("отправится", true) }
    val discrepancies = all.filter { it.hasDiscrepancyWithOtmeny }

    Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
        Text("Сводная отчетность по станции: $stationName", fontSize = 18.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
        Spacer(modifier = Modifier.height(10.dp))

        Row(horizontalArrangement = Arrangement.spacedBy(12.dp), modifier = Modifier.fillMaxWidth()) {
            StatSummaryBox("Отменено поездов", cancelled.size, TrainRTheme.RedCancel, Modifier.weight(1f))
            StatSummaryBox("Назначено поездов", appointed.size, TrainRTheme.GreenSuccess, Modifier.weight(1f))
            StatSummaryBox("Изменено время отпр.", shifted.size, TrainRTheme.GoldAccent, Modifier.weight(1f))
            StatSummaryBox("Расхождений с выгрузкой", discrepancies.size, TrainRTheme.NavyDark, Modifier.weight(1f))
        }

        Spacer(modifier = Modifier.height(14.dp))

        // Готовая строка сводки точно по примеру пользователя: "Фирсановская: отменено 5 поездов, 7001 Тверь-Москва в 16:01..."
        Surface(
            modifier = Modifier.fillMaxWidth(),
            shape = RoundedCornerShape(8.dp),
            color = Color(0xFFF8FAFC),
            border = androidx.compose.foundation.BorderStroke(1.dp, TrainRTheme.BorderCol)
        ) {
            Column(modifier = Modifier.padding(12.dp)) {
                val cancSummary = if (cancelled.isEmpty()) "$stationName: отменено 0 поездов."
                else "$stationName: отменено ${cancelled.size} поездов — " + cancelled.joinToString(", ") {
                    "№${it.trainNum} ${it.routeString} в ${if (it.afishaTime != "-") it.afishaTime else it.telegramTime}"
                }
                Text(cancSummary, fontSize = 12.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.RedCancel)

                Spacer(modifier = Modifier.height(6.dp))

                val appSummary = if (appointed.isEmpty()) "$stationName: назначено 0 поездов."
                else "$stationName: назначено ${appointed.size} поездов — " + appointed.joinToString(", ") {
                    "№${it.trainNum} ${it.routeString} в ${it.telegramTime}"
                }
                Text(appSummary, fontSize = 12.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.GreenSuccess)
            }
        }

        Spacer(modifier = Modifier.height(12.dp))

        LazyColumn(modifier = Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(6.dp)) {
            if (cancelled.isNotEmpty()) {
                item { Text("Детальный список отмененных поездов по ст. $stationName (${cancelled.size}):", fontWeight = FontWeight.Black, fontSize = 13.sp, color = TrainRTheme.RedCancel) }
                items(cancelled) { r ->
                    ReportTrainLine(r, TrainRTheme.RedCancel) { onTrainClick(r) }
                }
            }
            if (appointed.isNotEmpty()) {
                item { Spacer(modifier = Modifier.height(8.dp)); Text("Детальный список назначенных поездов и остановок (${appointed.size}):", fontWeight = FontWeight.Black, fontSize = 13.sp, color = TrainRTheme.GreenSuccess) }
                items(appointed) { r ->
                    ReportTrainLine(r, TrainRTheme.GreenSuccess) { onTrainClick(r) }
                }
            }
            if (discrepancies.isNotEmpty()) {
                item { Spacer(modifier = Modifier.height(8.dp)); Text("Расхождения между Телеграммой и файлом «Отмены» (${discrepancies.size}):", fontWeight = FontWeight.Black, fontSize = 13.sp, color = TrainRTheme.AmberWarn) }
                items(discrepancies) { r ->
                    ReportTrainLine(r, TrainRTheme.AmberWarn) { onTrainClick(r) }
                }
            }
        }
    }
}

@Composable
fun StatSummaryBox(label: String, count: Int, color: Color, modifier: Modifier) {
    Surface(
        modifier = modifier,
        shape = RoundedCornerShape(8.dp),
        color = color.copy(alpha = 0.08f),
        border = androidx.compose.foundation.BorderStroke(1.dp, color.copy(alpha = 0.3f))
    ) {
        Row(modifier = Modifier.padding(12.dp), verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.SpaceBetween) {
            Text(label, fontSize = 12.sp, fontWeight = FontWeight.ExtraBold, color = TrainRTheme.NavyDark)
            Text(count.toString(), fontSize = 20.sp, fontWeight = FontWeight.Black, color = color)
        }
    }
}

@Composable
fun ReportTrainLine(r: StationComparisonRow, accent: Color, onClick: () -> Unit) {
    Row(
        modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF8FAFC)).clickable { onClick() }.padding(8.dp),
        verticalAlignment = Alignment.CenterVertically,
        horizontalArrangement = Arrangement.SpaceBetween
    ) {
        Text("№${r.trainNum} (${r.routeString}) — в ${if (r.afishaTime != "-") r.afishaTime else r.telegramTime} [${r.dateGroup}]", fontWeight = FontWeight.ExtraBold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
        Text(r.finalChangeText, fontWeight = FontWeight.Black, fontSize = 12.sp, color = accent)
    }
}

// ==================== МОДАЛЬНОЕ ДОСЬЕ КОНКРЕТНОГО ПОЕЗДА ====================

@Composable
fun TrainDossierDialog(row: StationComparisonRow, onDismiss: () -> Unit) {
    Dialog(onDismissRequest = onDismiss) {
        Card(
            modifier = Modifier.width(760.dp).heightIn(min = 420.dp, max = 640.dp),
            shape = RoundedCornerShape(12.dp),
            colors = CardDefaults.cardColors(containerColor = Color.White),
            elevation = CardDefaults.cardElevation(8.dp)
        ) {
            Column(modifier = Modifier.fillMaxSize().padding(18.dp)) {
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    Column {
                        Text("ПОЛНЫЙ ОТЧЕТ ПО ПОЕЗДУ №${row.trainNum} ${row.daysFlag}", fontSize = 18.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark)
                        Text("Маршрут: ${row.routeString} (${row.destinationLabel}) • Источник: ${row.telegramHeader}", fontSize = 12.sp, fontWeight = FontWeight.Bold, color = TrainRTheme.RoyalBlue)
                    }
                    IconButton(onClick = onDismiss) {
                        Icon(Icons.Default.Close, contentDescription = null)
                    }
                }

                Divider(modifier = Modifier.padding(vertical = 8.dp))

                Text("Период действия: ${row.dateGroup}", fontWeight = FontWeight.ExtraBold, fontSize = 13.sp, color = TrainRTheme.NavyDark)
                Text("Итоговый статус по ст. ${row.stationName}: ${row.finalChangeText}", fontWeight = FontWeight.Black, fontSize = 13.sp, color = if (row.isCancelledAtStation) TrainRTheme.RedCancel else TrainRTheme.GreenSuccess)

                if (row.skippedStopsSummary.isNotEmpty()) {
                    Text("Отменены остановки: ${row.skippedStopsSummary}", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = TrainRTheme.RedCancel)
                }
                if (row.appointedStopsSummary.isNotEmpty()) {
                    Text("Назначены остановки: ${row.appointedStopsSummary}", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = TrainRTheme.GreenSuccess)
                }

                Spacer(modifier = Modifier.height(10.dp))
                Text("Сравнение расписания по всем станциям следования:", fontWeight = FontWeight.ExtraBold, fontSize = 12.sp, color = TrainRTheme.NavyDark)
                Spacer(modifier = Modifier.height(4.dp))

                val allStKeys = (row.fullAfishaStations.keys + row.fullTelegramStations.keys).distinct()
                LazyColumn(modifier = Modifier.weight(1f).fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF8FAFC)).padding(8.dp)) {
                    items(allStKeys) { stKey ->
                        val afT = row.fullAfishaStations[stKey] ?: "-"
                        val telPair = row.fullTelegramStations[stKey]
                        val telT = if (telPair != null) {
                            if (telPair.second) telPair.first else "${telPair.first} (без ост.)"
                        } else "-"

                        Row(modifier = Modifier.fillMaxWidth().padding(vertical = 3.dp), horizontalArrangement = Arrangement.SpaceBetween) {
                            Text(displayStationName(stKey), fontWeight = FontWeight.Bold, fontSize = 12.sp, modifier = Modifier.width(220.dp))
                            Text("По Афише: $afT", fontFamily = FontFamily.Monospace, fontSize = 12.sp, fontWeight = FontWeight.SemiBold, modifier = Modifier.width(160.dp))
                            Text(
                                text = "По Телеграмме: $telT",
                                fontFamily = FontFamily.Monospace,
                                fontSize = 12.sp,
                                fontWeight = FontWeight.ExtraBold,
                                color = if (telT != "-" && telT != afT) TrainRTheme.RoyalBlue else TrainRTheme.TextSub
                            )
                        }
                        Divider(color = Color(0xFFE2E8F0))
                    }
                }

                Spacer(modifier = Modifier.height(8.dp))
                Text("Исходный текст из телеграммы:", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = TrainRTheme.TextSub)
                Surface(modifier = Modifier.fillMaxWidth().height(70.dp), shape = RoundedCornerShape(6.dp), color = Color(0xFFF1F5F9)) {
                    Text(row.rawTelegramText, fontSize = 10.sp, modifier = Modifier.padding(6.dp), color = TrainRTheme.NavyDark)
                }
            }
        }
    }
}

// ==================== ВСПОМОГАТЕЛЬНЫЕ ЭЛЕМЕНТЫ UI ====================

@Composable
fun CompactFileBadge(
    title: String,
    subtitle: String,
    isLoaded: Boolean,
    accent: Color,
    modifier: Modifier,
    onSelect: () -> Unit,
    onClear: () -> Unit
) {
    Surface(
        modifier = modifier.height(50.dp).clickable { onSelect() },
        shape = RoundedCornerShape(8.dp),
        color = if (isLoaded) accent.copy(alpha = 0.08f) else Color(0xFFF8FAFC),
        border = androidx.compose.foundation.BorderStroke(1.dp, if (isLoaded) accent else TrainRTheme.BorderCol)
    ) {
        Row(modifier = Modifier.fillMaxSize().padding(horizontal = 8.dp), verticalAlignment = Alignment.CenterVertically) {
            Icon(
                imageVector = if (isLoaded) Icons.Default.CheckCircle else Icons.Default.Add,
                contentDescription = null,
                tint = if (isLoaded) accent else TrainRTheme.TextSub,
                modifier = Modifier.size(18.dp)
            )
            Spacer(modifier = Modifier.width(6.dp))
            Column(modifier = Modifier.weight(1f)) {
                Text(title, fontSize = 10.sp, fontWeight = FontWeight.Black, color = TrainRTheme.NavyDark, maxLines = 1)
                Text(subtitle, fontSize = 10.sp, fontWeight = FontWeight.SemiBold, color = TrainRTheme.TextSub, maxLines = 1, overflow = TextOverflow.Ellipsis)
            }
            if (isLoaded) {
                IconButton(onClick = onClear, modifier = Modifier.size(22.dp)) {
                    Icon(Icons.Default.Clear, contentDescription = "Удалить", tint = TrainRTheme.RedCancel, modifier = Modifier.size(15.dp))
                }
            }
        }
    }
}

@Composable
fun StationChipItem(
    name: String,
    isChecked: Boolean,
    isSelected: Boolean,
    modifier: Modifier,
    onCheck: () -> Unit,
    onSelect: () -> Unit
) {
    Surface(
        modifier = modifier.height(26.dp).clickable { onSelect() },
        shape = RoundedCornerShape(5.dp),
        color = if (isSelected) TrainRTheme.RoyalBlue else Color(0xFFF8FAFC),
        border = androidx.compose.foundation.BorderStroke(1.dp, if (isSelected) TrainRTheme.RoyalBlue else TrainRTheme.BorderCol)
    ) {
        Row(modifier = Modifier.fillMaxSize().padding(horizontal = 4.dp), verticalAlignment = Alignment.CenterVertically) {
            Checkbox(
                checked = isChecked,
                onCheckedChange = { onCheck() },
                modifier = Modifier.size(16.dp),
                colors = CheckboxDefaults.colors(
                    checkedColor = if (isSelected) Color.White else TrainRTheme.RoyalBlue,
                    checkmarkColor = if (isSelected) TrainRTheme.RoyalBlue else Color.White
                )
            )
            Spacer(modifier = Modifier.width(3.dp))
            Text(
                text = name.replace("Зеленоград-", "").replace(" ГРЭС", ""),
                fontSize = 10.sp,
                fontWeight = FontWeight.ExtraBold,
                color = if (isSelected) Color.White else TrainRTheme.NavyDark,
                maxLines = 1,
                overflow = TextOverflow.Ellipsis
            )
        }
    }
}

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "Train-R — Интеллектуальный ревизор пригородных расписаний (Тверь – Крюково – Москва – Ипподром МЦД)",
        state = WindowState(width = 1360.dp, height = 860.dp)
    ) {
        TrainRApp()
    }
}
