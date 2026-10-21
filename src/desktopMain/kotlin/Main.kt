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
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.window.Window
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

data class MasterTrainRecord(
    val dateGroup: String,
    val trainNum: String,
    val daysInfo: String,
    val station: String,
    val destination: String,
    val route: String,
    val afishaTime: String,
    val telegramTime: String,
    val vygruzkaTime: String,
    val telStatusText: String,
    val vygStatusText: String,
    val isCancelled: Boolean,
    val isAppointed: Boolean,
    val hasDiscrepancy: Boolean,
    val telegramHeader: String,
    val direction: Direction
)

enum class Direction { TO_MOSCOW, FROM_MOSCOW }

object AppTheme {
    val Background = Color(0xFFF8FAFC)
    val CardBg = Color(0xFFFFFFFF)
    val HeaderNavy = Color(0xFF0F172A)
    val CorporateBlue = Color(0xFF1E3A8A)
    val SuccessGreen = Color(0xFF00B96B) 
    val DangerRed = Color(0xFFF73B3B) 
    val WarningOrange = Color(0xFFF97316)
    val BorderLight = Color(0xFFE2E8F0)
    val TextPrimary = Color(0xFF0F172A)
    val TextMuted = Color(0xFF475569)
}

val ALL_STATIONS = listOf(
    "Алабушево", "Молжаниново", "Зеленоград-Крюково", "Химки", "Сходня", "Фирсановская",
    "Подрезково", "Новоподрезково", "Левобережная", "Ховрино", "Грачёвская",
    "Моссельмаш", "Лихоборы", "Останкино", "Подсолнечная", "Клин", "Тверь"
)

fun getCachedAfishaFile(): File {
    val dir = File(System.getProperty("user.home"), ".schedule_verifier")
    if (!dir.exists()) dir.mkdirs()
    return File(dir, "saved_afisha.xlsx")
}

fun matchStationNames(target: String, candidate: String): Boolean {
    val t = target.lowercase().replace("ё", "е").replace(" ", "").replace("-", "").replace(".", "")
    val c = candidate.lowercase().replace("ё", "е").replace(" ", "").replace("-", "").replace(".", "")
    if (t.isEmpty() || c.isEmpty()) return false
    if (t in c || c in t) return true

    val aliasMap = mapOf(
        "левобереж" to listOf("левобережн", "левобережье"),
        "крюков" to listOf("крюково", "зеленоград"),
        "петр" to listOf("петр", "петровско", "разумовск"),
        "покровк" to listOf("покровка"),
        "фирсановск" to listOf("фирсановская"),
        "подрезков" to listOf("подрезково"),
        "новоподрезков" to listOf("новоподрезково"),
        "молжанинов" to listOf("молжаниново"),
        "химк" to listOf("химки"),
        "сходн" to listOf("сходня"),
        "ховрин" to listOf("ховрино"),
        "грачевск" to listOf("грачевская", "грачевск"),
        "моссельмаш" to listOf("моссельмаш"),
        "лихобор" to listOf("лихоборы"),
        "останкин" to listOf("останкино"),
        "москв" to listOf("москва", "москвапасс", "москпасс", "ипподром"),
        "алабушев" to listOf("алабушево"),
        "радищев" to listOf("радищево"),
        "поваровк" to listOf("поваровка"),
        "поварово" to listOf("поварово"),
        "березк" to listOf("березки"),
        "подсолнечн" to listOf("подсолнечная"),
        "клин" to listOf("клин"),
        "твер" to listOf("тверь")
    )
    for ((stem, variants) in aliasMap) {
        if (stem in t) {
            for (v in variants) {
                if (v in c) return true
            }
        }
    }
    return false
}

fun cleanTime(raw: String): String {
    if (raw.isBlank()) return ""
    val clean = raw.split(",")[0].trim().replace(".", ":")
    val parts = clean.split(":")
    if (parts.size >= 2) {
        val h = parts[0].padStart(2, '0')
        val m = parts[1].padStart(2, '0')
        return "$h:$m"
    }
    return clean
}

fun unpackDates(raw: String, defaultHeader: String): List<String> {
    val clean = raw.replace("(", "").replace(")", "").trim()
    val monthsMap = mapOf(
        "01" to "января", "02" to "февраля", "03" to "марта", "04" to "апреля", "05" to "мая", "06" to "июня",
        "07" to "июля", "08" to "августа", "09" to "сентября", "10" to "октября", "11" to "ноября", "12" to "декабря",
        "1" to "января", "2" to "февраля", "3" to "марта", "4" to "апреля", "5" to "мая", "6" to "июня",
        "7" to "июля", "8" to "августа", "9" to "сентября"
    )

    val rangeRegex = Regex("с\\s*(\\d{1,2})[/.\\-](\\d{1,2})\\s*по\\s*(\\d{1,2})[/.\\-](\\d{1,2})")
    val rangeMatch = rangeRegex.find(clean)
    if (rangeMatch != null) {
        val (d1, _, d2, m2) = rangeMatch.destructured
        val mName = monthsMap[m2] ?: "октября"
        var res = "с ${d1.toInt()} по ${d2.toInt()} $mName"
        val krMatch = Regex("кр\\.?\\s*(\\d+)[/.\\-]?(\\d+)?").find(clean)
        if (krMatch != null) res += " (кроме ${krMatch.groupValues[1].toInt()} $mName)"
        return listOf(res)
    }

    val dashRange = Regex("(\\d{1,2})[/.\\-](\\d{1,2})\\s*[-–—]\\s*(\\d{1,2})[/.\\-](\\d{1,2})")
    val dashMatch = dashRange.find(clean)
    if (dashMatch != null) {
        val (d1, _, d2, m2) = dashMatch.destructured
        val mName = monthsMap[m2] ?: "октября"
        return listOf("с ${d1.toInt()} по ${d2.toInt()} $mName")
    }

    val listRegex = Regex("([\\d\\s,]+)[/.\\-](\\d{1,2})")
    val listMatch = listRegex.find(clean)
    if (listMatch != null) {
        val days = listMatch.groupValues[1].split(",").mapNotNull { it.trim().toIntOrNull() }
        val mName = monthsMap[listMatch.groupValues[2]] ?: "октября"
        return days.map { "$it $mName" }
    }

    return listOf(if (clean.isNotEmpty()) clean else defaultHeader)
}

fun extractFirstDateNumber(dateHeader: String): Int {
    val m = Regex("\\d+").find(dateHeader)
    return m?.value?.toIntOrNull() ?: 999
}

data class ParsedTrainBlock(
    val trainNum: String,
    val daysInfo: String,
    val datesRaw: String,
    val fullScheduleText: String,
    val stationsMap: Map<String, Pair<String, Boolean>>,
    val isCancelled: Boolean,
    val isAppointed: Boolean,
    val passWithoutStopStations: List<String>,
    val appointedStops: List<String>,
    val routeDesc: String
)

fun parseTelegramDetailed(telText: String, defaultPeriod: String): List<ParsedTrainBlock> {
    val result = mutableListOf<ParsedTrainBlock>()
    val lines = telText.lines()
    val timeTokenRegex = "\\d{1,2}(?:[:.]\\d{2})(?:[:.]\\d{2})?(?:,\\d+)?"
    val stationSegRegex = Regex("([А-Яа-яЁё0-9\\-\\.\\s]+?)\\s+(?:отпр\\.?\\s*|приб\\.?\\s*)?($timeTokenRegex)(?:\\s*[-–—]\\s*($timeTokenRegex))?")

    var inCancelSection = false
    var inAppointSection = false
    var i = 0

    while (i < lines.size) {
        val line = lines[i].trim()
        if (line.isEmpty()) { i++; continue }

        if (Regex("(?i)(?:^|\\s|[0-9.]+|\\n)(?:отменить|не назначать|отменяются)\\s*:").containsMatchIn(line)) {
            inCancelSection = true; inAppointSection = false
        } else if (Regex("(?i)(?:^|\\s|[0-9.]+|\\n)назначить\\s*:").containsMatchIn(line)) {
            inCancelSection = false; inAppointSection = true
        } else if (Regex("(?i)(?:^|\\s|[0-9.]+|\\n)изменить\\s*:").containsMatchIn(line)) {
            inCancelSection = false; inAppointSection = false
        }

        val trainMatch = Regex("(?i)(?:нр\\.?|поезд\\s*№?)\\s*([0-9]{4}(?:/[0-9]{4})?)\\s*(СВ|Р)?").find(line)
        if (trainMatch != null) {
            val trainNum = trainMatch.groupValues[1]
            val daysInfo = trainMatch.groupValues[2].uppercase()

            val dateInParens = Regex("\\(([^)]+)\\)").find(line)?.groupValues?.get(1)
            val dateRange = Regex("(?:с\\s*[0-9/.]+\\s*по\\s*[0-9/.]+.*)").find(line)?.value
            val datesRaw = dateInParens ?: dateRange ?: defaultPeriod
            
            val trainLineText = line.substringBefore("(").substringBefore(" с ")
            val routeMatch = Regex("[А-Яа-яЁё]+\\s*(?:Пасс\\.?|Тов\\.?|ГРЭС|Моск\\.?|1|2)?\\s*[-–—]\\s*[А-Яа-яЁё]+\\s*(?:Пасс\\.?|Тов\\.?|ГРЭС|Моск\\.?|1|2)?").find(trainLineText)
            val parsedRoute = routeMatch?.value?.trim() ?: ""

            val isCancelled = inCancelSection || line.contains("отменен", true) || line.contains("отменить", true) || line.contains("не назначать", true)
            val isAppointed = inAppointSection

            val fullBlockBuilder = StringBuilder(line)
            var j = i + 1
            while (j < lines.size) {
                val nextLine = lines[j].trim()
                if (nextLine.isEmpty()) { j++; continue }
                if (Regex("(?i)(?:нр\\.?|поезд\\s*№?)\\s*[0-9]{4}").containsMatchIn(nextLine) ||
                    Regex("(?i)(?:отменить|изменить|назначить)\\s*:").containsMatchIn(nextLine)
                ) break
                fullBlockBuilder.append(" ").append(nextLine)
                j++
            }
            i = j - 1

            val fullBlock = fullBlockBuilder.toString()
            val stationsMap = mutableMapOf<String, Pair<String, Boolean>>()
            val passWithoutStopList = mutableListOf<String>()
            val appointedStopsList = mutableListOf<String>()

            val passNotes = Regex("(?i)(?:оп|рп|ст)?\\s*([А-Яа-яЁё0-9\\-\\.\\s,]+?)\\s*проследует\\s*без\\s*остановки").findAll(fullBlock)
            for (pn in passNotes) {
                pn.groupValues[1].split(",").forEach { nm ->
                    val cleanNm = nm.trim().lowercase()
                    if (cleanNm.length >= 3) passWithoutStopList.add(cleanNm)
                }
            }

            val appNotes = Regex("(?i)(?:оп|рп|ст)?\\s*([А-Яа-яЁё0-9\\-\\.\\s,]+?)\\s*назначается\\s*остановка").findAll(fullBlock)
            for (an in appNotes) {
                an.groupValues[1].split(",").forEach { nm ->
                    val cleanNm = nm.trim().lowercase()
                    if (cleanNm.length >= 3) appointedStopsList.add(cleanNm)
                }
            }

            val segments = fullBlock.split(",")
            for (seg in segments) {
                val s = seg.trim()
                if (s.isEmpty()) continue
                val m = stationSegRegex.find(s)
                if (m != null) {
                    val stName = m.groupValues[1].replace(Regex("(?i)(?:нр\\.?|поезд\\s*№?)\\s*[0-9/]+"), "").replace("отпр.", "").replace("приб.", "").trim()
                    val t1 = m.groupValues[2]
                    val t2 = if (m.groupValues.size > 3) m.groupValues[3] else ""
                    val hasRange = t2.isNotEmpty()
                    val hasDep = s.contains("отпр", true)
                    val hasArr = s.contains("приб", true)
                    val isTech = s.contains("тех", true)

                    val depTime = if (hasRange) cleanTime(t2) else cleanTime(t1)
                    val isStop = (hasRange || hasDep || hasArr) && !isTech

                    if (stName.length >= 3) stationsMap[stName.lowercase()] = Pair(depTime, isStop)
                }
            }

            result.add(ParsedTrainBlock(trainNum, daysInfo, datesRaw, fullBlock, stationsMap, isCancelled, isAppointed, passWithoutStopList, appointedStopsList, parsedRoute))
        }
        i++
    }
    return result
}

fun extractTelegramHeader(telText: String): String {
    val m = Regex("(?i)ТЕЛЕГРАММА\\s*№?\\s*([^\\s]+)\\s+от\\s+([0-9/.-]+)").find(telText)
    if (m != null) return "Тел. №${m.groupValues[1]} от ${m.groupValues[2]}"
    return "Телеграмма РЖД"
}

fun analyzeMasterSchedule(
    afisha: File?,
    telegrams: List<File>,
    vygruzki: List<File>,
    targetStation: String
): Pair<List<MasterTrainRecord>, List<MasterTrainRecord>> {
    val toMsk = mutableListOf<MasterTrainRecord>()
    val fromMsk = mutableListOf<MasterTrainRecord>()

    val telText = buildString {
        telegrams.forEach { f ->
            try {
                if (f.name.endsWith(".docx", true)) FileInputStream(f).use { fis -> XWPFDocument(fis).use { doc -> doc.paragraphs.forEach { append(it.text).append("\n") } } }
                else if (f.name.endsWith(".pdf", true)) Loader.loadPDF(f).use { doc -> append(PDFTextStripper().getText(doc)).append("\n") }
            } catch (_: Exception) {}
        }
    }

    val defaultPeriod = Regex("(?i)с\\s*(\\d{1,2}\\s+[а-я]+(?:\\s*\\d{4})?\\s*г?\\.?\\s*по\\s*\\d{1,2}\\s+[а-я]+.*)").find(telText)?.value ?: "Текущий график"
    val telHeader = extractTelegramHeader(telText)
    val parsedBlocks = parseTelegramDetailed(telText, defaultPeriod)

    val vygText = buildString {
        vygruzki.forEach { f -> try { append(f.readText()).append("\n") } catch (_: Exception) {} }
    }

    // Глубокий парсинг Афиши со всех столбцов
    val afishaTimes = mutableMapOf<String, String>()
    if (afisha != null) {
        try {
            FileInputStream(afisha).use { fis ->
                val wb = WorkbookFactory.create(fis)
                val timeFmt = SimpleDateFormat("HH:mm")
                for (sheetIndex in 0 until wb.numberOfSheets) {
                    val sheet = wb.getSheetAt(sheetIndex)
                    val trainRows = mutableListOf<Row>()
                    for (r in 0..Math.min(20, sheet.lastRowNum)) {
                        val row = sheet.getRow(r) ?: continue
                        for (c in 0..row.lastCellNum) {
                            val v = row.getCell(c)?.toString()?.trim() ?: ""
                            if (v.matches(Regex("\\d{4}(/\\d{4})?"))) { trainRows.add(row); break }
                        }
                    }
                    var stationRow: Row? = null
                    for (r in 0..sheet.lastRowNum) {
                        val name = sheet.getRow(r)?.getCell(0)?.toString()?.trim() ?: ""
                        if (matchStationNames(targetStation, name)) { stationRow = sheet.getRow(r); break }
                    }
                    if (trainRows.isNotEmpty() && stationRow != null) {
                        for (tRow in trainRows) {
                            for (c in 0..tRow.lastCellNum) {
                                val rawNum = tRow.getCell(c)?.toString()?.trim()?.replace(".0", "") ?: ""
                                if (rawNum.matches(Regex("\\d{4}(/\\d{4})?"))) {
                                    val cell = stationRow.getCell(c)
                                    var aTime = ""
                                    if (cell != null) {
                                        if (cell.cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) aTime = timeFmt.format(cell.dateCellValue)
                                        else aTime = cell.toString().trim()
                                    }
                                    aTime = cleanTime(aTime)
                                    if (aTime.isNotEmpty() && aTime != "-") afishaTimes[rawNum.split("/").first()] = aTime
                                }
                            }
                        }
                    }
                }
            }
        } catch (_: Exception) {}
    }

    val processedBlocks = mutableSetOf<String>()

    // Анализ Телеграмм (Связка с Афишей)
    parsedBlocks.forEach { block ->
        val unpackedDates = unpackDates(block.datesRaw, defaultPeriod)
        val numOnly = block.trainNum.split("/").first()
        
        // Строгое определение направления по ЧЕТНОСТИ (Четные - От Москвы, Нечетные - На Москву)
        val isEven = numOnly.lastOrNull { it.isDigit() }?.digitToInt()?.let { it % 2 == 0 } ?: false
        val dir = if (isEven) Direction.FROM_MOSCOW else Direction.TO_MOSCOW
        val dest = if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Ипподрома/Твери"

        val afTime = afishaTimes[numOnly] ?: "-"
        
        var vygStatusText = ""
        var vygTime = ""
        val vygRegex = Regex("№$numOnly.*?отправлением в ([0-9:]+)\\s*-\\s*(.*?)(?=\\n|\$)").find(vygText)
        if (vygRegex != null) {
            vygTime = vygRegex.groupValues[1]
            vygStatusText = vygRegex.groupValues[2].trim()
        }

        unpackedDates.forEach { dateHeader ->
            var telStatusText = ""
            var isPass = false
            var isApp = block.isAppointed
            var telTime = ""

            if (block.isCancelled) {
                telStatusText = "ОТМЕНЕН"
            } else {
                val explicitlySkipped = block.passWithoutStopStations.any { matchStationNames(targetStation, it) }
                val explicitlyAppointed = block.appointedStops.any { matchStationNames(targetStation, it) }
                val stKey = block.stationsMap.keys.firstOrNull { matchStationNames(targetStation, it) }

                if (explicitlyAppointed) {
                    isApp = true
                    telTime = block.stationsMap[stKey]?.first ?: afTime.replace("-", "")
                    val ext = if (telTime.isNotEmpty()) ", отправится в $telTime" else ""
                    telStatusText = "Назначена остановка о.п. $targetStation$ext"
                } else if (explicitlySkipped) {
                    isPass = true
                    telStatusText = "ОТМЕНЕН (проследует без остановки)"
                } else if (stKey != null) {
                    val (depTime, isStop) = block.stationsMap[stKey]!!
                    if (!isStop) {
                        isPass = true
                        telStatusText = "ОТМЕНЕН (проследует без остановки)"
                    } else if (depTime != afTime && depTime.isNotEmpty()) {
                        telTime = depTime
                        telStatusText = "отправится в $depTime"
                    }
                }
            }

            // Фильтрация ложных изменений, если Афиша загружена
            if (telStatusText == "отправится в $afTime") telStatusText = ""
            if (vygStatusText == "отправится в $afTime") vygStatusText = ""

            // Если есть реальные изменения, отмены или станция явно упоминается в телеграмме
            if (block.isCancelled || isPass || isApp || telStatusText.isNotEmpty() || (vygStatusText.isNotEmpty() && !vygStatusText.contains(telStatusText, true)) || block.stationsMap.keys.any { matchStationNames(targetStation, it) }) {
                
                if (block.isCancelled || isPass || isApp || telStatusText.isNotEmpty() || vygStatusText.isNotEmpty()) {
                    val discrepancy = vygStatusText.isNotEmpty() && !vygStatusText.contains(telStatusText, true) && telStatusText != "ОТМЕНЕН"
                    
                    val record = MasterTrainRecord(
                        dateGroup = dateHeader, trainNum = block.trainNum, daysInfo = block.daysInfo, station = targetStation, destination = dest, route = block.routeDesc,
                        afishaTime = afTime, telegramTime = telTime, vygruzkaTime = vygTime, telStatusText = telStatusText, vygStatusText = vygStatusText,
                        isCancelled = block.isCancelled, isAppointed = isApp, hasDiscrepancy = discrepancy, telegramHeader = telHeader, direction = dir
                    )
                    if (dir == Direction.TO_MOSCOW) toMsk.add(record) else fromMsk.add(record)
                    processedBlocks.add(numOnly)
                }
            }
        }
    }

    // Анализ Выгрузки txt (поезда, которых нет в телеграмме, но есть в txt)
    val vygTrainRegex = Regex("№([0-9]{4}(?:/[0-9]{4})?).*?отправлением в ([0-9:]+)\\s*-\\s*(.*?)(?=\\n|\$)")
    vygTrainRegex.findAll(vygText).forEach { match ->
        val fullNum = match.groupValues[1]
        val numOnly = fullNum.split("/").first()
        
        if (!processedBlocks.contains(numOnly)) {
            val isEven = numOnly.lastOrNull { it.isDigit() }?.digitToInt()?.let { it % 2 == 0 } ?: false
            val dir = if (isEven) Direction.FROM_MOSCOW else Direction.TO_MOSCOW
            val dest = if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Ипподрома/Твери"
            
            val vTime = match.groupValues[2]
            var vStat = match.groupValues[3].trim()
            val afTime = afishaTimes[numOnly] ?: "-"
            
            if (vStat == "отправится в $afTime") vStat = ""
            if (vStat.isNotEmpty()) {
                val record = MasterTrainRecord(
                    dateGroup = defaultPeriod, trainNum = fullNum, daysInfo = "", station = targetStation, destination = dest, route = "",
                    afishaTime = afTime, telegramTime = "", vygruzkaTime = vTime, telStatusText = "", vygStatusText = vStat,
                    isCancelled = vStat.contains("ОТМЕНЕН", true), isAppointed = false, hasDiscrepancy = true, telegramHeader = "Выгрузка txt", direction = dir
                )
                if (dir == Direction.TO_MOSCOW) toMsk.add(record) else fromMsk.add(record)
                processedBlocks.add(numOnly)
            }
        }
    }

    val sortedToMsk = toMsk.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.afishaTime}" }.sortedWith(compareBy({ extractFirstDateNumber(it.dateGroup) }, { it.afishaTime.ifEmpty { it.telegramTime.ifEmpty { it.vygruzkaTime } } }))
    val sortedFromMsk = fromMsk.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.afishaTime}" }.sortedWith(compareBy({ extractFirstDateNumber(it.dateGroup) }, { it.afishaTime.ifEmpty { it.telegramTime.ifEmpty { it.vygruzkaTime } } }))

    return Pair(sortedToMsk, sortedFromMsk)
}

fun exportMasterReports(station: String, toMoscow: List<MasterTrainRecord>, fromMoscow: List<MasterTrainRecord>) {
    val desktop = File(System.getProperty("user.home"), "Desktop")
    val dirToMsk = File(desktop, "На Москву").apply { if (!exists()) mkdirs() }
    val dirFromMsk = File(desktop, "От Москвы").apply { if (!exists()) mkdirs() }

    if (toMoscow.isNotEmpty()) {
        val fToMsk = File(dirToMsk, "$station на Москву.txt")
        fToMsk.printWriter(Charsets.UTF_8).use { out ->
            out.println("$station Изменения в графике движения эл/поездов")
            out.println("на Москву")
            val grouped = toMoscow.groupBy { it.dateGroup }
            grouped.keys.sortedBy { extractFirstDateNumber(it) }.forEach { dateGroup ->
                out.println(dateGroup)
                grouped[dateGroup]?.forEach { r ->
                    // Вырезаем "(проследует без остановки)" для чистой выгрузки
                    val status = r.telStatusText.ifEmpty { r.vygStatusText }.replace(" \\(проследует без остановки\\)".toRegex(), "")
                    val destStr = if (r.route.isNotEmpty()) r.route else "до Москвы"
                    val line = java.lang.StringBuilder("        №${r.trainNum} ${r.daysInfo} $destStr")

                    // Умный вывод: если Афиша не загружена, не дублируем ложное отправление
                    if (r.afishaTime != "-" && r.afishaTime.isNotEmpty()) {
                        line.append(" отправлением в ${r.afishaTime} - $status")
                    } else {
                        if (status.startsWith("отправится в")) {
                            line.append(" - $status")
                        } else {
                            val backupTime = r.telegramTime.ifEmpty { r.vygruzkaTime }
                            if (backupTime.isNotEmpty()) line.append(" отправлением в $backupTime")
                            line.append(" - $status")
                        }
                    }
                    out.println(line.toString())
                }
            }
        }
    }

    if (fromMoscow.isNotEmpty()) {
        val fFromMsk = File(dirFromMsk, "$station от Москвы.txt")
        fFromMsk.printWriter(Charsets.UTF_8).use { out ->
            out.println("$station Изменения в графике движения эл/поездов")
            out.println("от Ленинградского вокзала, Ипподрома")
            val grouped = fromMoscow.groupBy { it.dateGroup }
            grouped.keys.sortedBy { extractFirstDateNumber(it) }.forEach { dateGroup ->
                out.println(dateGroup)
                grouped[dateGroup]?.forEach { r ->
                    val status = r.telStatusText.ifEmpty { r.vygStatusText }.replace(" \\(проследует без остановки\\)".toRegex(), "")
                    val line = java.lang.StringBuilder("        №${r.trainNum} ${r.daysInfo}")

                    if (r.afishaTime != "-" && r.afishaTime.isNotEmpty()) {
                        line.append(" отправлением в ${r.afishaTime} - $status")
                    } else {
                        if (status.startsWith("отправится в")) {
                            line.append(" - $status")
                        } else {
                            val backupTime = r.telegramTime.ifEmpty { r.vygruzkaTime }
                            if (backupTime.isNotEmpty()) line.append(" отправлением в $backupTime")
                            line.append(" - $status")
                        }
                    }
                    out.println(line.toString())
                }
            }
        }
    }
}

@Composable
fun App() {
    val cachedFile = remember { getCachedAfishaFile() }
    var afishaFile by remember { mutableStateOf<File?>(if (cachedFile.exists()) cachedFile else null) }
    var telegramFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var vygruzkiFiles by remember { mutableStateOf<List<File>>(emptyList()) }

    var selectedStations by remember { mutableStateOf(setOf("Сходня", "Подрезково")) }
    var currentStationView by remember { mutableStateOf("Сходня") }
    var searchQuery by remember { mutableStateOf("") }

    var toMoscowList by remember { mutableStateOf<List<MasterTrainRecord>>(emptyList()) }
    var fromMoscowList by remember { mutableStateOf<List<MasterTrainRecord>>(emptyList()) }
    var activeTab by remember { mutableStateOf(0) } 

    var isProcessing by remember { mutableStateOf(false) }
    var statusMessage by remember { mutableStateOf(if (telegramFiles.isEmpty() && vygruzkiFiles.isEmpty()) "Загрузите Телеграммы или Выгрузки (Афиша опционально)." else "Готов к глобальной сверке.") }

    val coroutineScope = rememberCoroutineScope()

    fun pickFiles(title: String, multiple: Boolean = false): List<File> {
        val d = FileDialog(null as Frame?, title, FileDialog.LOAD).apply { isMultipleMode = multiple; isVisible = true }
        val res = d.files.toList()
        d.dispose()
        return res
    }

    fun runGlobalSearch() {
        if (searchQuery.isBlank() || (telegramFiles.isEmpty() && vygruzkiFiles.isEmpty())) return
        coroutineScope.launch {
            isProcessing = true
            statusMessage = "Глобальный поиск нитки ${searchQuery.trim()} по всему полигону..."
            val globalToMsk = mutableListOf<MasterTrainRecord>()
            val globalFromMsk = mutableListOf<MasterTrainRecord>()
            withContext(Dispatchers.IO) {
                ALL_STATIONS.forEach { st ->
                    val (t, f) = analyzeMasterSchedule(afishaFile, telegramFiles, vygruzkiFiles, st)
                    globalToMsk.addAll(t.filter { it.trainNum.contains(searchQuery.trim(), true) })
                    globalFromMsk.addAll(f.filter { it.trainNum.contains(searchQuery.trim(), true) })
                }
            }
            toMoscowList = globalToMsk.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.station}" }.sortedBy { extractFirstDateNumber(it.dateGroup) }
            fromMoscowList = globalFromMsk.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.station}" }.sortedBy { extractFirstDateNumber(it.dateGroup) }
            isProcessing = false
            statusMessage = "Глобальный поиск завершен. Результаты выведены."
        }
    }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = AppTheme.Background) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                
                Card(modifier = Modifier.fillMaxWidth(), shape = RoundedCornerShape(12.dp), colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg), elevation = CardDefaults.cardElevation(2.dp)) {
                    Row(modifier = Modifier.fillMaxWidth().padding(14.dp), verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.SpaceBetween) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Icon(Icons.Default.CheckCircle, contentDescription = null, tint = AppTheme.CorporateBlue, modifier = Modifier.size(28.dp))
                            Spacer(modifier = Modifier.width(10.dp))
                            Column {
                                Text("SCHEDULE VERIFIER PRO — ТРОЙНАЯ СВЕРКА", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.HeaderNavy)
                                Text(statusMessage, fontSize = 12.sp, color = AppTheme.TextMuted, fontWeight = FontWeight.Bold)
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                    DocUploadCard(
                        title = "1. БАЗОВАЯ АФИША (ОПЦИЯ)", desc = if (afishaFile != null) "✓ ${afishaFile!!.name}" else "Файл Excel (.xlsx)",
                        isLoaded = afishaFile != null, color = AppTheme.SuccessGreen, icon = Icons.Default.DateRange, modifier = Modifier.weight(1f),
                        onClick = {
                            val f = pickFiles("Выберите Афишу (.xlsx)")
                            if (f.isNotEmpty()) {
                                try { FileInputStream(f.first()).use { input -> FileOutputStream(cachedFile).use { output -> input.copyTo(output) } }; afishaFile = cachedFile; statusMessage = "Афиша обновлена!" } catch (_: Exception) { afishaFile = f.first() }
                            }
                        },
                        onClear = { if (cachedFile.exists()) cachedFile.delete(); afishaFile = null; statusMessage = "Афиша удалена из памяти." }
                    )

                    DocUploadCard(
                        title = "2. ТЕЛЕГРАММЫ", desc = if (telegramFiles.isNotEmpty()) "Выбрано: ${telegramFiles.size} док." else "Word / PDF",
                        isLoaded = telegramFiles.isNotEmpty(), color = AppTheme.CorporateBlue, icon = Icons.Default.Info, modifier = Modifier.weight(1f),
                        onClick = { val f = pickFiles("Выберите Телеграммы", true); if (f.isNotEmpty()) telegramFiles = f; statusMessage = "Телеграммы загружены." },
                        onClear = { telegramFiles = emptyList(); statusMessage = "Телеграммы удалены." }
                    )

                    DocUploadCard(
                        title = "3. ВЫГРУЗКИ (ОПЦИЯ)", desc = if (vygruzkiFiles.isNotEmpty()) "Выбрано: ${vygruzkiFiles.size} док." else "Текстовые файлы (.txt)",
                        isLoaded = vygruzkiFiles.isNotEmpty(), color = Color(0xFF6366F1), icon = Icons.Default.List, modifier = Modifier.weight(1f),
                        onClick = { val f = pickFiles("Выберите Выгрузки (.txt)", true); if (f.isNotEmpty()) vygruzkiFiles = f; statusMessage = "Выгрузки загружены." },
                        onClear = { vygruzkiFiles = emptyList(); statusMessage = "Выгрузки удалены." }
                    )
                }

                Spacer(modifier = Modifier.height(10.dp))

                Card(modifier = Modifier.fillMaxWidth(), shape = RoundedCornerShape(10.dp), colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg), elevation = CardDefaults.cardElevation(1.dp)) {
                    Column(modifier = Modifier.fillMaxWidth().padding(12.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text("Станции полигона:", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = AppTheme.HeaderNavy)
                            Spacer(modifier = Modifier.width(8.dp))
                            FilledTonalButton(onClick = { selectedStations = if (selectedStations.size == ALL_STATIONS.size) setOf("Сходня") else ALL_STATIONS.toSet() }, contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp), shape = RoundedCornerShape(6.dp)) {
                                Icon(Icons.Default.Done, contentDescription = null, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text(if (selectedStations.size == ALL_STATIONS.size) "Снять все" else "Выбрать все станции сразу", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }
                            Spacer(modifier = Modifier.weight(1f))
                            Button(
                                onClick = {
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Тройная сверка по ст. $currentStationView..."
                                        val (toMsk, fromMsk) = withContext(Dispatchers.IO) { analyzeMasterSchedule(afishaFile, telegramFiles, vygruzkiFiles, currentStationView) }
                                        toMoscowList = toMsk; fromMoscowList = fromMsk; isProcessing = false
                                        statusMessage = "Сверка ст. $currentStationView завершена!"
                                    }
                                },
                                enabled = !isProcessing && (telegramFiles.isNotEmpty() || vygruzkiFiles.isNotEmpty()), colors = ButtonDefaults.buttonColors(containerColor = AppTheme.CorporateBlue), shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(16.dp)); Spacer(modifier = Modifier.width(6.dp)); Text(if (isProcessing) "Сверить $currentStationView" else "Сверить $currentStationView")
                            }
                            Spacer(modifier = Modifier.width(8.dp))
                            Button(
                                onClick = {
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Генерация отчетов по папкам 'На Москву' и 'От Москвы'..."
                                        withContext(Dispatchers.IO) { selectedStations.forEach { st -> val (toMsk, fromMsk) = analyzeMasterSchedule(afishaFile, telegramFiles, vygruzkiFiles, st); exportMasterReports(st, toMsk, fromMsk) } }
                                        isProcessing = false; statusMessage = "УСПЕШНО! Файлы отмен разложены по папкам 'На Москву' и 'От Москвы' на Рабочем столе!"
                                    }
                                },
                                enabled = !isProcessing && selectedStations.isNotEmpty() && (telegramFiles.isNotEmpty() || vygruzkiFiles.isNotEmpty()), colors = ButtonDefaults.buttonColors(containerColor = AppTheme.DangerRed), shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Share, contentDescription = null, modifier = Modifier.size(16.dp)); Spacer(modifier = Modifier.width(6.dp)); Text("Создать файлы ОТМЕНА в папки (${selectedStations.size} ст.)")
                            }
                        }
                        Spacer(modifier = Modifier.height(6.dp))
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) { ALL_STATIONS.take(8).forEach { st -> CompactStationChip(st, selectedStations.contains(st), currentStationView == st, onToggle = { selectedStations = if (selectedStations.contains(st)) selectedStations - st else selectedStations + st }, onSelectView = { currentStationView = st }) } }
                        Spacer(modifier = Modifier.height(4.dp))
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) { ALL_STATIONS.drop(8).forEach { st -> CompactStationChip(st, selectedStations.contains(st), currentStationView == st, onToggle = { selectedStations = if (selectedStations.contains(st)) selectedStations - st else selectedStations + st }, onSelectView = { currentStationView = st }) } }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    TabRow(selectedTabIndex = activeTab, containerColor = AppTheme.CardBg, contentColor = AppTheme.CorporateBlue, modifier = Modifier.width(550.dp).clip(RoundedCornerShape(8.dp))) {
                        Tab(selected = (activeTab == 0), onClick = { activeTab = 0 }, text = { Text("НА МОСКВУ (${toMoscowList.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp) })
                        Tab(selected = (activeTab == 1), onClick = { activeTab = 1 }, text = { Text("ОТ МОСКВЫ (${fromMoscowList.size})", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp) })
                        Tab(selected = (activeTab == 2), onClick = { activeTab = 2 }, text = { Text("ОТЧЕТЫ (СТАТИСТИКА)", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.SuccessGreen) })
                    }

                    Row(verticalAlignment = Alignment.CenterVertically) {
                        OutlinedTextField(
                            value = searchQuery, onValueChange = { searchQuery = it }, placeholder = { Text("Поиск нитки...", fontSize = 12.sp, fontWeight = FontWeight.Bold) },
                            modifier = Modifier.width(200.dp).height(50.dp), singleLine = true, textStyle = LocalTextStyle.current.copy(fontSize = 13.sp, fontWeight = FontWeight.Bold), shape = RoundedCornerShape(8.dp)
                        )
                        Spacer(modifier = Modifier.width(6.dp))
                        Button(onClick = { runGlobalSearch() }, colors = ButtonDefaults.buttonColors(containerColor = AppTheme.CorporateBlue), shape = RoundedCornerShape(8.dp), modifier = Modifier.height(50.dp)) {
                            Icon(Icons.Default.Search, contentDescription = null, modifier = Modifier.size(18.dp)); Spacer(modifier = Modifier.width(4.dp)); Text("Найти", fontWeight = FontWeight.Bold)
                        }
                    }
                }

                Spacer(modifier = Modifier.height(8.dp))

                Card(modifier = Modifier.fillMaxSize(), shape = RoundedCornerShape(10.dp), colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg), elevation = CardDefaults.cardElevation(2.dp)) {
                    if (activeTab == 2) {
                        val totalList = toMoscowList + fromMoscowList
                        val cancList = totalList.filter { it.telStatusText.contains("ОТМЕНЕН", true) }
                        val appList = totalList.filter { it.telStatusText.contains("Назначен", true) }
                        val modList = totalList.filter { it.telStatusText.contains("отправится", true) && !it.telStatusText.contains("Назначен", true) }
                        val discrList = totalList.filter { it.hasDiscrepancy }
                        
                        Column(modifier = Modifier.fillMaxSize().padding(20.dp)) {
                            Text("Сводный аналитический отчет по станции: $currentStationView", fontSize = 18.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.HeaderNavy)
                            Spacer(modifier = Modifier.height(24.dp))
                            
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Canvas(modifier = Modifier.size(14.dp)) { drawCircle(color = AppTheme.DangerRed) }
                                Spacer(modifier = Modifier.width(10.dp))
                                Text("Отменено поездов: ${cancList.size}", fontSize = 18.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.DangerRed)
                            }
                            Spacer(modifier = Modifier.height(12.dp))
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Canvas(modifier = Modifier.size(14.dp)) { drawCircle(color = AppTheme.SuccessGreen) }
                                Spacer(modifier = Modifier.width(10.dp))
                                Text("Назначено поездов: ${appList.size}", fontSize = 18.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.SuccessGreen)
                            }
                            Spacer(modifier = Modifier.height(12.dp))
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Canvas(modifier = Modifier.size(14.dp)) { drawCircle(color = AppTheme.WarningOrange) }
                                Spacer(modifier = Modifier.width(10.dp))
                                Text("Изменено времени отправления: ${modList.size}", fontSize = 16.sp, fontWeight = FontWeight.Bold, color = AppTheme.WarningOrange)
                            }
                            Spacer(modifier = Modifier.height(12.dp))
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Canvas(modifier = Modifier.size(14.dp)) { drawCircle(color = AppTheme.HeaderNavy) }
                                Spacer(modifier = Modifier.width(10.dp))
                                Text("Расхождений с выгрузками txt: ${discrList.size}", fontSize = 16.sp, fontWeight = FontWeight.Bold, color = AppTheme.HeaderNavy)
                            }

                            Spacer(modifier = Modifier.height(24.dp))
                            Divider(color = AppTheme.BorderLight)
                            Spacer(modifier = Modifier.height(12.dp))
                            
                            LazyColumn(modifier = Modifier.fillMaxWidth().weight(1f)) {
                                if (cancList.isNotEmpty()) {
                                    item { Text("Отмененные поезда:", fontWeight = FontWeight.ExtraBold, fontSize = 14.sp, color = AppTheme.DangerRed, modifier = Modifier.padding(bottom = 6.dp)) }
                                    item { Text(cancList.joinToString(", ") { "№${it.trainNum} (${it.afishaTime.ifEmpty { it.telegramTime }}) ${it.destination}" }, fontSize = 13.sp, fontWeight = FontWeight.Bold, color = AppTheme.TextPrimary, modifier = Modifier.padding(bottom = 12.dp)) }
                                }
                                if (appList.isNotEmpty()) {
                                    item { Text("Назначенные поезда:", fontWeight = FontWeight.ExtraBold, fontSize = 14.sp, color = AppTheme.SuccessGreen, modifier = Modifier.padding(bottom = 6.dp)) }
                                    item { Text(appList.joinToString(", ") { "№${it.trainNum} (${it.telegramTime}) ${it.destination}" }, fontSize = 13.sp, fontWeight = FontWeight.Bold, color = AppTheme.TextPrimary, modifier = Modifier.padding(bottom = 12.dp)) }
                                }
                                if (discrList.isNotEmpty()) {
                                    item { Text("Расхождения с выгрузками:", fontWeight = FontWeight.ExtraBold, fontSize = 14.sp, color = AppTheme.WarningOrange, modifier = Modifier.padding(bottom = 6.dp)) }
                                    item { Text(discrList.joinToString(", ") { "№${it.trainNum} (${it.afishaTime.ifEmpty { it.telegramTime }}) ${it.destination}" }, fontSize = 13.sp, fontWeight = FontWeight.Bold, color = AppTheme.TextPrimary, modifier = Modifier.padding(bottom = 12.dp)) }
                                }
                            }
                        }
                    } else {
                        val currentList = (if (activeTab == 0) toMoscowList else fromMoscowList).filter { it.trainNum.contains(searchQuery.trim(), ignoreCase = true) }
                        Column(modifier = Modifier.fillMaxSize().padding(10.dp)) {
                            Row(modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF1F5F9)).padding(8.dp), verticalAlignment = Alignment.CenterVertically) {
                                Text("ДАТЫ", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(120.dp))
                                Text("ПОЕЗД И ДОК.", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(100.dp))
                                Text("МАРШРУТ", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(130.dp))
                                Text("АФИША", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(60.dp))
                                Text("СТАТУС ТЕЛЕГРАММЫ", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.weight(1f))
                                Text("ВЫГРУЗКА (.TXT)", fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(150.dp))
                            }
                            Divider(color = AppTheme.BorderLight, modifier = Modifier.padding(vertical = 4.dp))
                            if (currentList.isEmpty()) { Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) { Text("Нажмите «Сверить» или «Найти».", color = AppTheme.TextMuted, fontWeight = FontWeight.Bold) } }
                            else {
                                LazyColumn(modifier = Modifier.fillMaxSize()) {
                                    items(currentList) { item ->
                                        val isCancel = item.telStatusText.contains("ОТМЕНЕН", true)
                                        val bg = if (item.hasDiscrepancy) Color(0xFFFEF3C7) else if (isCancel) Color(0xFFFFF5F5) else Color.White
                                        val border = if (item.hasDiscrepancy) Color(0xFFFDE68A) else if (isCancel) Color(0xFFFEE2E2) else Color(0xFFF1F5F9)

                                        Row(modifier = Modifier.fillMaxWidth().padding(vertical = 2.dp).clip(RoundedCornerShape(6.dp)).background(bg).border(1.dp, border, RoundedCornerShape(6.dp)).padding(vertical = 6.dp, horizontal = 8.dp), verticalAlignment = Alignment.CenterVertically) {
                                            Text(item.dateGroup, fontWeight = FontWeight.ExtraBold, fontSize = 11.sp, color = AppTheme.CorporateBlue, modifier = Modifier.width(120.dp))
                                            Column(modifier = Modifier.width(100.dp)) {
                                                Text("№${item.trainNum} ${item.daysInfo}", fontWeight = FontWeight.ExtraBold, fontFamily = FontFamily.Monospace, fontSize = 12.sp, color = AppTheme.HeaderNavy)
                                                Text(item.telegramHeader, fontSize = 9.sp, color = AppTheme.TextMuted, fontWeight = FontWeight.Bold, maxLines = 1)
                                            }
                                            Text(item.route.ifEmpty { "-" }, fontSize = 10.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.CorporateBlue, modifier = Modifier.width(130.dp), maxLines = 2)
                                            Text(item.afishaTime, fontFamily = FontFamily.Monospace, fontSize = 12.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.TextMuted, modifier = Modifier.width(60.dp))
                                            Text(item.telStatusText, fontSize = 11.sp, fontWeight = if (isCancel || item.isAppointed) FontWeight.ExtraBold else FontWeight.Bold, color = if (isCancel) AppTheme.DangerRed else if (item.telStatusText.contains("Назначен", true)) AppTheme.SuccessGreen else AppTheme.TextPrimary, modifier = Modifier.weight(1f))
                                            Text(item.vygStatusText, fontSize = 10.sp, fontWeight = FontWeight.Bold, color = if (item.hasDiscrepancy) AppTheme.DangerRed else AppTheme.TextMuted, modifier = Modifier.width(150.dp), maxLines = 2)
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

@Composable
fun CompactStationChip(name: String, isChecked: Boolean, isCurrentView: Boolean, onToggle: () -> Unit, onSelectView: () -> Unit) {
    Surface(shape = RoundedCornerShape(6.dp), color = if (isCurrentView) AppTheme.CorporateBlue.copy(alpha = 0.12f) else Color(0xFFF8FAFC), border = androidx.compose.foundation.BorderStroke(1.dp, if (isCurrentView) AppTheme.CorporateBlue else AppTheme.BorderLight), modifier = Modifier.clickable { onSelectView() }) {
        Row(modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp), verticalAlignment = Alignment.CenterVertically) {
            Checkbox(checked = isChecked, onCheckedChange = { onToggle() }, modifier = Modifier.size(24.dp)); Spacer(modifier = Modifier.width(2.dp))
            Text(name.replace("Зеленоград-", ""), fontSize = 11.sp, fontWeight = if (isCurrentView) FontWeight.ExtraBold else FontWeight.Bold, color = if (isCurrentView) AppTheme.CorporateBlue else AppTheme.TextPrimary)
        }
    }
}

@Composable
fun DocUploadCard(title: String, desc: String, isLoaded: Boolean, color: Color, icon: ImageVector, modifier: Modifier, onClick: () -> Unit, onClear: () -> Unit) {
    Card(modifier = modifier.clickable { onClick() }, shape = RoundedCornerShape(10.dp), colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg), elevation = CardDefaults.cardElevation(1.dp)) {
        Row(modifier = Modifier.fillMaxWidth().padding(12.dp), verticalAlignment = Alignment.CenterVertically) {
            Surface(shape = RoundedCornerShape(8.dp), color = if (isLoaded) color.copy(alpha = 0.12f) else Color(0xFFF1F5F9), modifier = Modifier.size(38.dp)) {
                Box(contentAlignment = Alignment.Center) { Icon(if (isLoaded) Icons.Default.Check else icon, contentDescription = null, tint = if (isLoaded) color else AppTheme.TextMuted, modifier = Modifier.size(20.dp)) }
            }
            Spacer(modifier = Modifier.width(10.dp))
            Column(modifier = Modifier.weight(1f)) {
                Text(title, fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.TextMuted)
                Text(desc, fontSize = 12.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.HeaderNavy, maxLines = 1)
            }
            if (isLoaded) { IconButton(onClick = onClear, modifier = Modifier.size(24.dp)) { Icon(Icons.Default.Clear, contentDescription = "Clear", tint = AppTheme.DangerRed, modifier = Modifier.size(18.dp)) } }
            else { TextButton(onClick = onClick) { Text("Выбрать", fontSize = 11.sp, fontWeight = FontWeight.ExtraBold, color = color) } }
        }
    }
}

fun main() = application { Window(onCloseRequest = ::exitApplication, title = "ScheduleVerifier PRO — Экспертная система") { App() } }
