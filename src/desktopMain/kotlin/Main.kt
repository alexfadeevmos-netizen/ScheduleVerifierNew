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

data class DatedTrainRecord(
    val dateGroup: String,
    val trainNum: String,
    val station: String,
    val destination: String,
    val afishaTime: String,
    val actualTime: String?,
    val isCancelled: Boolean,
    val isPassWithoutStop: Boolean,
    val changeText: String,
    val direction: Direction
)

enum class Direction { TO_MOSCOW, FROM_MOSCOW }

object AppTheme {
    val Background = Color(0xFFF8FAFC)
    val CardBg = Color(0xFFFFFFFF)
    val HeaderNavy = Color(0xFF0F172A)
    val CorporateBlue = Color(0xFF1E3A8A)
    val SuccessGreen = Color(0xFF059669)
    val DangerRed = Color(0xFFDC2626)
    val BorderLight = Color(0xFFE2E8F0)
    val TextPrimary = Color(0xFF1E293B)
    val TextMuted = Color(0xFF64748B)
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
        "москв" to listOf("москва", "москвапасс", "москпасс"),
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

// Преобразование дат в читаемый заголовок
fun humanizeDateString(raw: String, defaultHeader: String): String {
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
        if ("кр" in clean) {
            val krMatch = Regex("кр\\.?\\s*(\\d+)[/.\\-]?(\\d+)?").find(clean)
            if (krMatch != null) {
                res += " (кроме ${krMatch.groupValues[1].toInt()} $mName)"
            }
        }
        return res
    }

    val dashRange = Regex("(\\d{1,2})[/.\\-](\\d{1,2})\\s*[-–—]\\s*(\\d{1,2})[/.\\-](\\d{1,2})")
    val dashMatch = dashRange.find(clean)
    if (dashMatch != null) {
        val (d1, _, d2, m2) = dashMatch.destructured
        val mName = monthsMap[m2] ?: "октября"
        return "с ${d1.toInt()} по ${d2.toInt()} $mName"
    }

    val listRegex = Regex("([\\d\\s,]+)[/.\\-](\\d{1,2})")
    val listMatch = listRegex.find(clean)
    if (listMatch != null) {
        val days = listMatch.groupValues[1].split(",").mapNotNull { it.trim().toIntOrNull() }
        val mName = monthsMap[listMatch.groupValues[2]] ?: "октября"
        return "${days.joinToString(", ")} $mName"
    }

    return if (clean.isNotEmpty()) clean else defaultHeader
}

// Извлечение первого числа даты для хронологической сортировки
fun extractFirstDateNumber(dateHeader: String): Int {
    val m = Regex("\\d+").find(dateHeader)
    return m?.value?.toIntOrNull() ?: 999
}

data class ParsedTrainBlock(
    val trainNum: String,
    val datesRaw: String,
    val fullScheduleText: String,
    val stationsMap: Map<String, Pair<String, Boolean>>,
    val isCancelled: Boolean,
    val passWithoutStopStations: List<String>
)

fun parseTelegramDetailed(telText: String, defaultPeriod: String): List<ParsedTrainBlock> {
    val result = mutableListOf<ParsedTrainBlock>()
    val lines = telText.lines()
    val timeTokenRegex = "\\d{1,2}(?:[:.]\\d{2})(?:[:.]\\d{2})?(?:,\\d+)?"
    val stationSegRegex = Regex("([А-Яа-яЁё0-9\\-\\.\\s]+?)\\s+(?:отпр\\.?\\s*|приб\\.?\\s*)?($timeTokenRegex)(?:\\s*[-–—]\\s*($timeTokenRegex))?")

    var inCancelSection = false
    var i = 0

    while (i < lines.size) {
        val line = lines[i].trim()
        if (line.isEmpty()) {
            i++
            continue
        }

        if (Regex("(?:^|\\s)(?:отменить|не назначать)\\s*:", RegexOption.IGNORE_CASE).containsMatchIn(line)) {
            inCancelSection = true
        } else if (Regex("(?:^|\\s)(?:изменить|назначить)\\s*:", RegexOption.IGNORE_CASE).containsMatchIn(line)) {
            inCancelSection = false
        }

        val trainMatch = Regex("(?:нр\\.?|поезд\\s*№?)\\s*([0-9]{4}(?:/[0-9]{4})?)", RegexOption.IGNORE_CASE).find(line)
        if (trainMatch != null) {
            val trainNum = trainMatch.groupValues[1]

            val dateInParens = Regex("\\(([^)]+)\\)").find(line)?.groupValues?.get(1)
            val dateRange = Regex("(?:с\\s*[0-9/.]+\\s*по\\s*[0-9/.]+.*)").find(line)?.value
            val datesRaw = dateInParens ?: dateRange ?: defaultPeriod

            val isCancelled = inCancelSection || line.contains("отменен", true) || line.contains("отменить", true) || line.contains("не назначать", true)

            val fullBlockBuilder = StringBuilder(line)
            var j = i + 1
            while (j < lines.size) {
                val nextLine = lines[j].trim()
                if (nextLine.isEmpty()) {
                    j++
                    continue
                }
                if (Regex("(?:нр\\.?|поезд\\s*№?)\\s*[0-9]{4}", RegexOption.IGNORE_CASE).containsMatchIn(nextLine) ||
                    Regex("(?:отменить|изменить|назначить)\\s*:", RegexOption.IGNORE_CASE).containsMatchIn(nextLine)
                ) {
                    break
                }
                fullBlockBuilder.append(" ").append(nextLine)
                j++
            }
            i = j - 1

            val fullBlock = fullBlockBuilder.toString()
            val stationsMap = mutableMapOf<String, Pair<String, Boolean>>()
            val passWithoutStopList = mutableListOf<String>()

            val passNotes = Regex("(?i)(?:оп|рп)?\\s*([А-Яа-яЁё0-9\\-\\.\\s,]+?)\\s*проследует\\s*без\\s*остановки").findAll(fullBlock)
            for (pn in passNotes) {
                val names = pn.groupValues[1].split(",")
                for (nm in names) {
                    val cleanNm = nm.trim().lowercase()
                    if (cleanNm.length >= 3) passWithoutStopList.add(cleanNm)
                }
            }

            val segments = fullBlock.split(",")
            for (seg in segments) {
                val s = seg.trim()
                if (s.isEmpty()) continue

                val m = stationSegRegex.find(s)
                if (m != null) {
                    val stName = m.groupValues[1]
                        .replace(Regex("(?i)(?:нр\\.?|поезд\\s*№?)\\s*[0-9/]+"), "")
                        .replace("отпр.", "")
                        .replace("приб.", "")
                        .trim()

                    val t1 = m.groupValues[2]
                    val t2 = if (m.groupValues.size > 3) m.groupValues[3] else ""

                    val hasRange = t2.isNotEmpty()
                    val hasDep = s.contains("отпр", true)
                    val hasArr = s.contains("приб", true)
                    val isTech = s.contains("тех", true)

                    val depTime = if (hasRange) cleanTime(t2) else cleanTime(t1)
                    val isStop = (hasRange || hasDep || hasArr) && !isTech

                    if (stName.length >= 3) {
                        stationsMap[stName.lowercase()] = Pair(depTime, isStop)
                    }
                }
            }

            result.add(
                ParsedTrainBlock(
                    trainNum = trainNum,
                    datesRaw = datesRaw,
                    fullScheduleText = fullBlock,
                    stationsMap = stationsMap,
                    isCancelled = isCancelled,
                    passWithoutStopStations = passWithoutStopList
                )
            )
        }
        i++
    }
    return result
}

fun extractDefaultTelegramPeriod(telText: String): String {
    val headerDate = Regex("(?i)с\\s*(\\d{1,2}\\s+[а-я]+(?:\\s*\\d{4})?\\s*г?\\.?\\s*по\\s*\\d{1,2}\\s+[а-я]+(?:\\s*\\d{4})?\\s*г?\\.?)").find(telText)?.value
    if (!headerDate.isNullOrEmpty()) return headerDate

    val singleHeader = Regex("(?i)с\\s*(\\d{1,2}\\s+[а-я]+(?:\\s*\\d{4})?\\s*г?\\.?)").find(telText)?.value
    if (!singleHeader.isNullOrEmpty()) return singleHeader

    return "с 3 по 8 октября"
}

// Полный анализ расписания со строгим разделением листов и хронологической сортировкой
fun analyzeScheduleWithSmartDates(
    afisha: File,
    telegrams: List<File>,
    vygruzki: List<File>,
    targetStation: String
): Pair<List<DatedTrainRecord>, List<DatedTrainRecord>> {
    val toMsk = mutableListOf<DatedTrainRecord>()
    val fromMsk = mutableListOf<DatedTrainRecord>()

    val telText = buildString {
        telegrams.forEach { f ->
            try {
                if (f.name.endsWith(".docx", true)) {
                    FileInputStream(f).use { fis -> XWPFDocument(fis).use { doc -> doc.paragraphs.forEach { append(it.text).append("\n") } } }
                } else if (f.name.endsWith(".pdf", true)) {
                    Loader.loadPDF(f).use { doc -> append(PDFTextStripper().getText(doc)).append("\n") }
                }
            } catch (_: Exception) {}
        }
    }

    val defaultPeriod = extractDefaultTelegramPeriod(telText)
    val parsedBlocks = parseTelegramDetailed(telText, defaultPeriod)

    val vygText = buildString {
        vygruzki.forEach { f ->
            try { append(f.readText()).append("\n") } catch (_: Exception) {}
        }
    }

    try {
        FileInputStream(afisha).use { fis ->
            val wb = WorkbookFactory.create(fis)
            val timeFmt = SimpleDateFormat("HH:mm")

            for (sheetIndex in 0 until wb.numberOfSheets) {
                val sheet = wb.getSheetAt(sheetIndex)
                val sheetName = sheet.sheetName.lowercase().trim()

                // Строгое определение направления
                val isFromMoscow = sheetName.startsWith("москва")
                val isToMoscow = !isFromMoscow && (sheetName.endsWith("москва") || sheetName.contains("на москву") || sheetName.contains("москва"))

                if (!isToMoscow && !isFromMoscow) continue
                val dir = if (isToMoscow) Direction.TO_MOSCOW else Direction.FROM_MOSCOW

                var trainRowIdx = -1
                for (r in 0..6) {
                    val row = sheet.getRow(r) ?: continue
                    for (c in 1..25) {
                        val v = row.getCell(c)?.toString()?.trim() ?: ""
                        if (v.matches(Regex("\\d{4}(/\\d{4})?"))) {
                            trainRowIdx = r
                            break
                        }
                    }
                    if (trainRowIdx != -1) break
                }
                if (trainRowIdx == -1) continue

                var stationRowIdx = -1
                for (r in trainRowIdx + 1..sheet.lastRowNum) {
                    val row = sheet.getRow(r) ?: continue
                    val name = row.getCell(0)?.toString()?.trim() ?: ""
                    if (matchStationNames(targetStation, name)) {
                        stationRowIdx = r
                        break
                    }
                }
                if (stationRowIdx == -1) continue

                val trainRow = sheet.getRow(trainRowIdx)
                val stRow = sheet.getRow(stationRowIdx)

                for (c in 1..trainRow.lastCellNum) {
                    val rawNum = trainRow.getCell(c.toInt())?.toString()?.trim()?.replace(".0", "") ?: ""
                    if (!rawNum.matches(Regex("\\d{4}(/\\d{4})?"))) continue

                    val cell = stRow.getCell(c.toInt())
                    var afishaTime = ""
                    if (cell != null) {
                        if (cell.cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                            afishaTime = timeFmt.format(cell.dateCellValue)
                        } else {
                            afishaTime = cell.toString().trim()
                        }
                    }
                    if (afishaTime.isEmpty() || afishaTime == "-") continue
                    afishaTime = cleanTime(afishaTime)

                    val numOnly = rawNum.split("/").first()
                    val matchingBlocks = parsedBlocks.filter { it.trainNum.split("/").first() == numOnly }

                    if (matchingBlocks.isNotEmpty()) {
                        matchingBlocks.forEach { block ->
                            val dateHeader = humanizeDateString(block.datesRaw, defaultPeriod)
                            var isCancelled = block.isCancelled
                            var isPassWithoutStop = false
                            var newTime: String? = null
                            var changeText = ""

                            if (isCancelled) {
                                changeText = "ОТМЕНЕН"
                            } else {
                                val explicitlySkipped = block.passWithoutStopStations.any { matchStationNames(targetStation, it) }
                                val stKey = block.stationsMap.keys.firstOrNull { matchStationNames(targetStation, it) }

                                if (explicitlySkipped) {
                                    isPassWithoutStop = true
                                    changeText = "ОТМЕНЕН (проследует без остановки)"
                                } else if (stKey != null) {
                                    val (depTime, isStop) = block.stationsMap[stKey]!!
                                    if (!isStop) {
                                        isPassWithoutStop = true
                                        changeText = "ОТМЕНЕН (проследует без остановки)"
                                    } else if (depTime != afishaTime && depTime.isNotEmpty()) {
                                        newTime = depTime
                                        changeText = "отправится в $depTime"
                                    }
                                }
                            }

                            if (isCancelled || isPassWithoutStop || newTime != null || changeText.isNotEmpty()) {
                                val dest = if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Крюково/Твери"
                                val record = DatedTrainRecord(
                                    dateGroup = dateHeader,
                                    trainNum = rawNum,
                                    station = targetStation,
                                    destination = dest,
                                    afishaTime = afishaTime,
                                    actualTime = newTime,
                                    isCancelled = isCancelled,
                                    isPassWithoutStop = isPassWithoutStop,
                                    changeText = changeText,
                                    direction = dir
                                )
                                if (dir == Direction.TO_MOSCOW) toMsk.add(record) else fromMsk.add(record)
                            }
                        }
                    } else if (vygText.isNotEmpty()) {
                        val vygRegex = Regex("№$numOnly.*?отправлением в [0-9:]+\\s*-\\s*(.*?)(?=\\n|\$)")
                        val vygMatch = vygRegex.find(vygText)
                        if (vygMatch != null) {
                            val desc = vygMatch.groupValues[1].trim()
                            val isCanc = desc.contains("ОТМЕНЕН", true)
                            val isPass = desc.contains("без остановки", true)
                            val dest = if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Крюково/Твери"
                            val record = DatedTrainRecord(
                                dateGroup = defaultPeriod,
                                trainNum = rawNum,
                                station = targetStation,
                                destination = dest,
                                afishaTime = afishaTime,
                                actualTime = null,
                                isCancelled = isCanc,
                                isPassWithoutStop = isPass,
                                changeText = desc,
                                direction = dir
                            )
                            if (dir == Direction.TO_MOSCOW) toMsk.add(record) else fromMsk.add(record)
                        }
                    }
                }
            }
        }
    } catch (_: Exception) {}

    // Устранение дубликатов и сортировка: по дням месяца, затем по времени отправления
    val sortedToMsk = toMsk.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.afishaTime}" }
        .sortedWith(compareBy({ extractFirstDateNumber(it.dateGroup) }, { it.afishaTime }))

    val sortedFromMsk = fromMsk.distinctBy { "${it.dateGroup}_${it.trainNum}_${it.afishaTime}" }
        .sortedWith(compareBy({ extractFirstDateNumber(it.dateGroup) }, { it.afishaTime }))

    return Pair(sortedToMsk, sortedFromMsk)
}

// Экспорт двух файлов строго в хронологическом порядке
fun exportDatedReports(station: String, toMoscow: List<DatedTrainRecord>, fromMoscow: List<DatedTrainRecord>) {
    val desktop = File(System.getProperty("user.home"), "Desktop")

    val fToMsk = File(desktop, "$station на Москву.txt")
    fToMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station Изменения в графике движения эл/поездов")
        out.println("на Москву")

        val grouped = toMoscow.groupBy { it.dateGroup }
        val sortedKeys = grouped.keys.sortedBy { extractFirstDateNumber(it) }

        sortedKeys.forEach { dateGroup ->
            out.println(dateGroup)
            grouped[dateGroup]?.forEach { r ->
                out.println("        №${r.trainNum} ${r.destination} отправлением в ${r.afishaTime} - ${r.changeText}")
            }
        }
    }

    val fFromMsk = File(desktop, "$station от Москвы.txt")
    fFromMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station Изменения в графике движения эл/поездов")
        out.println("от Ленинградского вокзала, Ипподрома")

        val grouped = fromMoscow.groupBy { it.dateGroup }
        val sortedKeys = grouped.keys.sortedBy { extractFirstDateNumber(it) }

        sortedKeys.forEach { dateGroup ->
            out.println(dateGroup)
            grouped[dateGroup]?.forEach { r ->
                out.println("        №${r.trainNum} отправлением в ${r.afishaTime} - ${r.changeText}")
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

    var selectedStations by remember { mutableStateOf(setOf("Сходня", "Подрезково", "Химки")) }
    var currentStationView by remember { mutableStateOf("Сходня") }
    var searchQuery by remember { mutableStateOf("") }

    var toMoscowList by remember { mutableStateOf<List<DatedTrainRecord>>(emptyList()) }
    var fromMoscowList by remember { mutableStateOf<List<DatedTrainRecord>>(emptyList()) }
    var activeTab by remember { mutableStateOf(0) }

    var isProcessing by remember { mutableStateOf(false) }
    var statusMessage by remember {
        mutableStateOf(
            if (afishaFile != null) "Базовая афиша активна. Нажмите «Сверить», чтобы увидеть отмены и сдвиги графика строго по хронологии дат."
            else "Загрузите файл базовой афиши расписания (.xlsx)."
        )
    }

    val coroutineScope = rememberCoroutineScope()

    fun pickFiles(title: String, multiple: Boolean = false): List<File> {
        val d = FileDialog(null as Frame?, title, FileDialog.LOAD).apply {
            isMultipleMode = multiple
            isVisible = true
        }
        val res = d.files.toList()
        d.dispose()
        return res
    }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = AppTheme.Background) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {

                // ШАПКА
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(14.dp),
                        verticalAlignment = Alignment.CenterVertically,
                        horizontalArrangement = Arrangement.SpaceBetween
                    ) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Icon(Icons.Default.CheckCircle, contentDescription = null, tint = AppTheme.CorporateBlue, modifier = Modifier.size(28.dp))
                            Spacer(modifier = Modifier.width(10.dp))
                            Column {
                                Text("SCHEDULE VERIFIER PRO — ХРОНОЛОГИЧЕСКИЙ РЕВИЗОР РАСПИСАНИЙ", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.HeaderNavy)
                                Text(statusMessage, fontSize = 12.sp, color = AppTheme.TextMuted)
                            }
                        }

                        if (toMoscowList.isNotEmpty() || fromMoscowList.isNotEmpty()) {
                            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFEFF6FF)) {
                                    Text("На Москву: ${toMoscowList.size}", modifier = Modifier.padding(horizontal = 10.dp, vertical = 6.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = AppTheme.CorporateBlue)
                                }
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFF5F3FF)) {
                                    Text("От Москвы: ${fromMoscowList.size}", modifier = Modifier.padding(horizontal = 10.dp, vertical = 6.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = Color(0xFF6D28D9))
                                }
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // КАРТОЧКИ ЗАГРУЗКИ
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                    DocUploadCard(
                        title = "1. БАЗОВАЯ АФИША (ПОСТОЯННАЯ)",
                        desc = if (afishaFile != null) "✓ Активна: ${afishaFile!!.name}" else "Выберите файл Excel (.xlsx)",
                        isLoaded = afishaFile != null,
                        color = AppTheme.SuccessGreen,
                        icon = Icons.Default.DateRange,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Афишу (.xlsx)")
                        if (f.isNotEmpty()) {
                            val selected = f.first()
                            try {
                                FileInputStream(selected).use { input ->
                                    FileOutputStream(cachedFile).use { output -> input.copyTo(output) }
                                }
                                afishaFile = cachedFile
                                statusMessage = "Афиша сохранена в постоянную память приложения!"
                            } catch (_: Exception) {
                                afishaFile = selected
                            }
                        }
                    }

                    DocUploadCard(
                        title = "2. ТЕЛЕГРАММЫ (ОСНОВА СВЕРКИ)",
                        desc = if (telegramFiles.isNotEmpty()) "Выбрано документов: ${telegramFiles.size}" else "Word (.docx) / PDF",
                        isLoaded = telegramFiles.isNotEmpty(),
                        color = AppTheme.CorporateBlue,
                        icon = Icons.Default.Info,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Телеграммы", multiple = true)
                        if (f.isNotEmpty()) {
                            telegramFiles = f
                            statusMessage = "Загружено телеграмм: ${f.size}"
                        }
                    }

                    DocUploadCard(
                        title = "3. ОПЕРАТИВНЫЕ ВЫГРУЗКИ (ОПЦИЯ)",
                        desc = if (vygruzkiFiles.isNotEmpty()) "Выбрано файлов: ${vygruzkiFiles.size}" else "Текстовые файлы (.txt)",
                        isLoaded = vygruzkiFiles.isNotEmpty(),
                        color = Color(0xFF6366F1),
                        icon = Icons.Default.List,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Выгрузки (.txt)", multiple = true)
                        if (f.isNotEmpty()) {
                            vygruzkiFiles = f
                            statusMessage = "Загружено выгрузок: ${f.size}"
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // ВЫБОР СТАНЦИЙ
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg),
                    elevation = CardDefaults.cardElevation(1.dp)
                ) {
                    Column(modifier = Modifier.fillMaxWidth().padding(12.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text("Станции полигона:", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = AppTheme.HeaderNavy)
                            Spacer(modifier = Modifier.width(8.dp))

                            // Кнопка ВЫБРАТЬ ВСЕ / СНЯТЬ ВСЕ
                            FilledTonalButton(
                                onClick = {
                                    selectedStations = if (selectedStations.size == ALL_STATIONS.size) setOf("Сходня") else ALL_STATIONS.toSet()
                                },
                                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp),
                                shape = RoundedCornerShape(6.dp)
                            ) {
                                Icon(Icons.Default.Done, contentDescription = null, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text(
                                    text = if (selectedStations.size == ALL_STATIONS.size) "Снять все" else "Выбрать все станции сразу",
                                    fontSize = 11.sp,
                                    fontWeight = FontWeight.Bold
                                )
                            }

                            Spacer(modifier = Modifier.weight(1f))

                            Button(
                                onClick = {
                                    if (afishaFile == null) return@Button
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Анализ расписания по ст. $currentStationView в обоих направлениях..."
                                        val (toMsk, fromMsk) = withContext(Dispatchers.IO) {
                                            analyzeScheduleWithSmartDates(afishaFile!!, telegramFiles, vygruzkiFiles, currentStationView)
                                        }
                                        toMoscowList = toMsk
                                        fromMoscowList = fromMsk
                                        isProcessing = false
                                        statusMessage = "Сверка ст. $currentStationView завершена! На Москву: ${toMsk.size}, От Москвы: ${fromMsk.size}"
                                    }
                                },
                                enabled = !isProcessing && afishaFile != null,
                                colors = ButtonDefaults.buttonColors(containerColor = AppTheme.CorporateBlue),
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text(if (isProcessing) "Сверка..." else "Сверить $currentStationView")
                            }

                            Spacer(modifier = Modifier.width(8.dp))

                            Button(
                                onClick = {
                                    if (afishaFile == null) return@Button
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Генерация отчетов по ${selectedStations.size} станциям на Рабочий стол..."
                                        withContext(Dispatchers.IO) {
                                            selectedStations.forEach { st ->
                                                val (toMsk, fromMsk) = analyzeScheduleWithSmartDates(afishaFile!!, telegramFiles, vygruzkiFiles, st)
                                                exportDatedReports(st, toMsk, fromMsk)
                                            }
                                        }
                                        isProcessing = false
                                        statusMessage = "УСПЕШНО! Созданы файлы отмен по ${selectedStations.size} станциям на Рабочем столе!"
                                    }
                                },
                                enabled = !isProcessing && afishaFile != null && selectedStations.isNotEmpty(),
                                colors = ButtonDefaults.buttonColors(containerColor = AppTheme.DangerRed),
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Share, contentDescription = null, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("Создать файлы ОТМЕНА (${selectedStations.size} ст.)")
                            }
                        }

                        Spacer(modifier = Modifier.height(6.dp))

                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                            ALL_STATIONS.take(8).forEach { st ->
                                CompactStationChip(st, selectedStations.contains(st), currentStationView == st,
                                    onToggle = { selectedStations = if (selectedStations.contains(st)) selectedStations - st else selectedStations + st },
                                    onSelectView = { currentStationView = st }
                                )
                            }
                        }
                        Spacer(modifier = Modifier.height(4.dp))
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                            ALL_STATIONS.drop(8).forEach { st ->
                                CompactStationChip(st, selectedStations.contains(st), currentStationView == st,
                                    onToggle = { selectedStations = if (selectedStations.contains(st)) selectedStations - st else selectedStations + st },
                                    onSelectView = { currentStationView = st }
                                )
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // ТАБЫ И ПОЛЕ ПОИСКА (БЕЗ ОБРЕЗАНИЯ ШРИФТА)
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    TabRow(
                        selectedTabIndex = activeTab,
                        containerColor = AppTheme.CardBg,
                        contentColor = AppTheme.CorporateBlue,
                        modifier = Modifier.width(440.dp).clip(RoundedCornerShape(8.dp))
                    ) {
                        Tab(
                            selected = (activeTab == 0),
                            onClick = { activeTab = 0 },
                            text = { Text("НА МОСКВУ (${toMoscowList.size})", fontWeight = FontWeight.Bold, fontSize = 12.sp) }
                        )
                        Tab(
                            selected = (activeTab == 1),
                            onClick = { activeTab = 1 },
                            text = { Text("ОТ МОСКВЫ (${fromMoscowList.size})", fontWeight = FontWeight.Bold, fontSize = 12.sp) }
                        )
                    }

                    OutlinedTextField(
                        value = searchQuery,
                        onValueChange = { searchQuery = it },
                        placeholder = { Text("Поиск нитки поезда (напр. 6405)...", fontSize = 12.sp) },
                        modifier = Modifier.width(320.dp).height(54.dp),
                        singleLine = true,
                        leadingIcon = { Icon(Icons.Default.Search, contentDescription = null, tint = AppTheme.TextMuted, modifier = Modifier.size(18.dp)) },
                        textStyle = LocalTextStyle.current.copy(fontSize = 13.sp, lineHeight = 18.sp),
                        shape = RoundedCornerShape(8.dp)
                    )
                }

                Spacer(modifier = Modifier.height(8.dp))

                val currentList = (if (activeTab == 0) toMoscowList else fromMoscowList)
                    .filter { it.trainNum.contains(searchQuery.trim(), ignoreCase = true) }

                Card(
                    modifier = Modifier.fillMaxSize(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Column(modifier = Modifier.fillMaxSize().padding(10.dp)) {
                        Row(
                            modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF1F5F9)).padding(8.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text("ПЕРИОД ДАТ (ПО ПОРЯДКУ)", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(170.dp))
                            Text("ПОЕЗД", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(90.dp))
                            Text("ПО АФИШЕ", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(100.dp))
                            Text("СТАТУС ИЗМЕНЕНИЙ ПО ТЕЛЕГРАММЕ", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.weight(1f))
                        }

                        Divider(color = AppTheme.BorderLight, modifier = Modifier.padding(vertical = 4.dp))

                        if (currentList.isEmpty()) {
                            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                Text("Нажмите «Сверить $currentStationView» для отображения результатов по датам.", color = AppTheme.TextMuted, fontSize = 13.sp)
                            }
                        } else {
                            LazyColumn(modifier = Modifier.fillMaxSize()) {
                                items(currentList) { item ->
                                    val isCancel = item.isCancelled || item.isPassWithoutStop
                                    val bg = if (isCancel) Color(0xFFFFF5F5) else Color.White
                                    val border = if (isCancel) Color(0xFFFEE2E2) else Color(0xFFF1F5F9)

                                    Row(
                                        modifier = Modifier
                                            .fillMaxWidth()
                                            .padding(vertical = 2.dp)
                                            .clip(RoundedCornerShape(6.dp))
                                            .background(bg)
                                            .border(1.dp, border, RoundedCornerShape(6.dp))
                                            .padding(vertical = 6.dp, horizontal = 8.dp),
                                        verticalAlignment = Alignment.CenterVertically
                                    ) {
                                        Text(
                                            text = item.dateGroup,
                                            fontWeight = FontWeight.Bold,
                                            fontSize = 12.sp,
                                            color = AppTheme.CorporateBlue,
                                            modifier = Modifier.width(170.dp)
                                        )

                                        Text(
                                            text = "№${item.trainNum}",
                                            fontWeight = FontWeight.Bold,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            color = AppTheme.HeaderNavy,
                                            modifier = Modifier.width(90.dp)
                                        )

                                        Text(
                                            text = item.afishaTime,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            fontWeight = FontWeight.SemiBold,
                                            color = AppTheme.TextMuted,
                                            modifier = Modifier.width(100.dp)
                                        )

                                        Text(
                                            text = item.changeText,
                                            fontSize = 12.sp,
                                            fontWeight = if (isCancel) FontWeight.Bold else FontWeight.Medium,
                                            color = if (isCancel) AppTheme.DangerRed else AppTheme.TextPrimary,
                                            modifier = Modifier.weight(1f)
                                        )
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
    Surface(
        shape = RoundedCornerShape(6.dp),
        color = if (isCurrentView) AppTheme.CorporateBlue.copy(alpha = 0.12f) else Color(0xFFF8FAFC),
        border = androidx.compose.foundation.BorderStroke(1.dp, if (isCurrentView) AppTheme.CorporateBlue else AppTheme.BorderLight),
        modifier = Modifier.clickable { onSelectView() }
    ) {
        Row(modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp), verticalAlignment = Alignment.CenterVertically) {
            Checkbox(
                checked = isChecked,
                onCheckedChange = { onToggle() },
                modifier = Modifier.size(24.dp)
            )
            Spacer(modifier = Modifier.width(2.dp))
            Text(
                text = name.replace("Зеленоград-", ""),
                fontSize = 11.sp,
                fontWeight = if (isCurrentView) FontWeight.Bold else FontWeight.Normal,
                color = if (isCurrentView) AppTheme.CorporateBlue else AppTheme.TextPrimary
            )
        }
    }
}

@Composable
fun DocUploadCard(title: String, desc: String, isLoaded: Boolean, color: Color, icon: ImageVector, modifier: Modifier, onClick: () -> Unit) {
    Card(
        modifier = modifier.clickable { onClick() },
        shape = RoundedCornerShape(10.dp),
        colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg),
        elevation = CardDefaults.cardElevation(1.dp)
    ) {
        Row(modifier = Modifier.fillMaxWidth().padding(12.dp), verticalAlignment = Alignment.CenterVertically) {
            Surface(shape = RoundedCornerShape(8.dp), color = if (isLoaded) color.copy(alpha = 0.12f) else Color(0xFFF1F5F9), modifier = Modifier.size(38.dp)) {
                Box(contentAlignment = Alignment.Center) {
                    Icon(if (isLoaded) Icons.Default.Check else icon, contentDescription = null, tint = if (isLoaded) color else AppTheme.TextMuted, modifier = Modifier.size(20.dp))
                }
            }
            Spacer(modifier = Modifier.width(10.dp))
            Column(modifier = Modifier.weight(1f)) {
                Text(title, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = AppTheme.TextMuted)
                Text(desc, fontSize = 12.sp, fontWeight = FontWeight.SemiBold, color = AppTheme.HeaderNavy, maxLines = 1)
            }
            TextButton(onClick = onClick) {
                Text(if (isLoaded) "Заменить" else "Выбрать", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = color)
            }
        }
    }
}

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "ScheduleVerifier PRO — Хронологический ревизор расписаний"
    ) {
        App()
    }
}
