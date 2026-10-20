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
import java.text.SimpleDateFormat

// Модель рейса с изменениями
data class TrainChangeRecord(
    val trainNum: String,
    val destination: String,
    val afishaTime: String,
    val newTime: String?,
    val isCancelled: Boolean,
    val changeText: String,
    val direction: Direction
)

enum class Direction {
    TO_MOSCOW, FROM_MOSCOW
}

object Theme {
    val Bg = Color(0xFFF8FAFC)
    val CardBg = Color(0xFFFFFFFF)
    val Navy = Color(0xFF0F172A)
    val AccentBlue = Color(0xFF1E3A8A)
    val Green = Color(0xFF059669)
    val Red = Color(0xFFDC2626)
    val Border = Color(0xFFE2E8F0)
    val TextDark = Color(0xFF1E293B)
    val TextMuted = Color(0xFF64748B)
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun App() {
    var afishaFile by remember { mutableStateOf<File?>(null) }
    var telegramFile by remember { mutableStateOf<File?>(null) }
    var vygruzkiFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var selectedStation by remember { mutableStateOf("Молжаниново") }
    var isExpandedDropdown by remember { mutableStateOf(false) }

    var toMoscowList by remember { mutableStateOf<List<TrainChangeRecord>>(emptyList()) }
    var fromMoscowList by remember { mutableStateOf<List<TrainChangeRecord>>(emptyList()) }
    var activeTab by remember { mutableStateOf(0) } // 0: На Москву, 1: От Москвы

    var isProcessing by remember { mutableStateOf(false) }
    var statusMessage by remember { mutableStateOf("Готов к работе. Загрузите Афишу, Телеграмму и файлы выгрузок.") }

    val coroutineScope = rememberCoroutineScope()
    val availableStations = listOf(
        "Молжаниново", "Зеленоград-Крюково", "Химки", "Сходня", "Фирсановская",
        "Подрезково", "Новоподрезково", "Левобережная", "Ховрино", "Грачёвская",
        "Моссельмаш", "Лихоборы", "Останкино", "Подсолнечная", "Клин", "Тверь"
    )

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
        Surface(modifier = Modifier.fillMaxSize(), color = Theme.Bg) {
            Column(modifier = Modifier.fillMaxSize().padding(18.dp)) {

                // ШАПКА
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = Theme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(16.dp),
                        verticalAlignment = Alignment.CenterVertically,
                        horizontalArrangement = Arrangement.SpaceBetween
                    ) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Icon(Icons.Default.CheckCircle, contentDescription = null, tint = Theme.AccentBlue, modifier = Modifier.size(28.dp))
                            Spacer(modifier = Modifier.width(10.dp))
                            Column {
                                Text("SCHEDULE VERIFIER PRO — ГЕНЕРАТОР ОТМЕН", fontSize = 17.sp, fontWeight = FontWeight.ExtraBold, color = Theme.Navy)
                                Text(statusMessage, fontSize = 13.sp, color = Theme.TextMuted)
                            }
                        }

                        if (toMoscowList.isNotEmpty() || fromMoscowList.isNotEmpty()) {
                            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFEFF6FF)) {
                                    Text("На Москву: ${toMoscowList.size}", modifier = Modifier.padding(8.dp, 4.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = Theme.AccentBlue)
                                }
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFF5F3FF)) {
                                    Text("От Москвы: ${fromMoscowList.size}", modifier = Modifier.padding(8.dp, 4.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = Color(0xFF6D28D9))
                                }
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))

                // КНОПКИ ЗАГРУЗКИ ФАЙЛОВ
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                    UploadCard(
                        title = "1. АФИША (.XLSX)",
                        desc = afishaFile?.name ?: "Выберите файл Excel афиши",
                        isLoaded = afishaFile != null,
                        color = Theme.Green,
                        icon = Icons.Default.DateRange,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Афишу (.xlsx)")
                        if (f.isNotEmpty()) {
                            afishaFile = f.first()
                            statusMessage = "Афиша: ${afishaFile?.name}"
                        }
                    }

                    UploadCard(
                        title = "2. ТЕЛЕГРАММА",
                        desc = telegramFile?.name ?: "Word (.docx) или PDF",
                        isLoaded = telegramFile != null,
                        color = Theme.AccentBlue,
                        icon = Icons.Default.Info,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Телеграмму (Word/PDF)")
                        if (f.isNotEmpty()) {
                            telegramFile = f.first()
                            statusMessage = "Телеграмма: ${telegramFile?.name}"
                        }
                    }

                    UploadCard(
                        title = "3. ВЫГРУЗКИ (.TXT)",
                        desc = if (vygruzkiFiles.isNotEmpty()) "Выбрано файлов: ${vygruzkiFiles.size}" else "Текстовые выгрузки",
                        isLoaded = vygruzkiFiles.isNotEmpty(),
                        color = Color(0xFF6366F1),
                        icon = Icons.Default.List,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Выгрузки (.txt)", multiple = true)
                        if (f.isNotEmpty()) {
                            vygruzkiFiles = f
                            statusMessage = "Выгрузок выбрано: ${f.size}"
                        }
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))

                // ПАНЕЛЬ ВЫБОРА СТАНЦИИ И ЭКСПОРТА
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = Theme.CardBg),
                    elevation = CardDefaults.cardElevation(1.dp)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(12.dp),
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Text("Станция:", fontWeight = FontWeight.Bold, fontSize = 13.sp, color = Theme.Navy)
                        Spacer(modifier = Modifier.width(8.dp))

                        // Выпадающий список выбора станции
                        ExposedDropdownMenuBox(
                            expanded = isExpandedDropdown,
                            onExpandedChange = { isExpandedDropdown = !isExpandedDropdown },
                            modifier = Modifier.width(220.dp)
                        ) {
                            OutlinedTextField(
                                value = selectedStation,
                                onValueChange = {},
                                readOnly = true,
                                trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = isExpandedDropdown) },
                                modifier = Modifier.menuAnchor().fillMaxWidth(),
                                textStyle = LocalTextStyle.current.copy(fontSize = 13.sp, fontWeight = FontWeight.SemiBold)
                            )
                            ExposedDropdownMenu(
                                expanded = isExpandedDropdown,
                                onDismissRequest = { isExpandedDropdown = false }
                            ) {
                                availableStations.forEach { station ->
                                    DropdownMenuItem(
                                        text = { Text(station, fontSize = 13.sp) },
                                        onClick = {
                                            selectedStation = station
                                            isExpandedDropdown = false
                                        }
                                    )
                                }
                            }
                        }

                        Spacer(modifier = Modifier.weight(1f))

                        // Кнопка ОБРАБОТАТЬ
                        Button(
                            onClick = {
                                if (afishaFile == null) return@Button
                                coroutineScope.launch {
                                    isProcessing = true
                                    statusMessage = "Формирование отчета по ст. $selectedStation..."
                                    val (toMsk, fromMsk) = withContext(Dispatchers.IO) {
                                        processStationSchedule(afishaFile!!, telegramFile, vygruzkiFiles, selectedStation)
                                    }
                                    toMoscowList = toMsk
                                    fromMoscowList = fromMsk
                                    isProcessing = false
                                    statusMessage = "Готово! Найдено изменений: на Москву (${toMsk.size}), от Москвы (${fromMsk.size})"
                                }
                            },
                            enabled = !isProcessing && afishaFile != null,
                            colors = ButtonDefaults.buttonColors(containerColor = Theme.AccentBlue),
                            shape = RoundedCornerShape(8.dp)
                        ) {
                            Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(16.dp))
                            Spacer(modifier = Modifier.width(6.dp))
                            Text(if (isProcessing) "Обработка..." else "Сверить станцию")
                        }

                        Spacer(modifier = Modifier.width(8.dp))

                        // Кнопка ЭКСПОРТ ДВУХ ФАЙЛОВ
                        Button(
                            onClick = {
                                if (toMoscowList.isEmpty() && fromMoscowList.isEmpty()) return@Button
                                exportStationReports(selectedStation, toMoscowList, fromMoscowList)
                                statusMessage = "УСПЕШНО! Созданы 2 файла: '$selectedStation на Москву.txt' и '$selectedStation от Москвы.txt' на Рабочем столе!"
                            },
                            enabled = toMoscowList.isNotEmpty() || fromMoscowList.isNotEmpty(),
                            colors = ButtonDefaults.buttonColors(containerColor = Theme.Red),
                            shape = RoundedCornerShape(8.dp)
                        ) {
                            Icon(Icons.Default.Share, contentDescription = null, modifier = Modifier.size(16.dp))
                            Spacer(modifier = Modifier.width(6.dp))
                            Text("Создать 2 файла (На Москву и От Москвы)")
                        }
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))

                // ТАБЫ: НА МОСКВУ / ОТ МОСКВЫ
                TabRow(
                    selectedTabIndex = activeTab,
                    containerColor = Theme.CardBg,
                    contentColor = Theme.AccentBlue,
                    modifier = Modifier.clip(RoundedCornerShape(8.dp))
                ) {
                    Tab(
                        selected = (activeTab == 0),
                        onClick = { activeTab = 0 },
                        text = { Text("НА МОСКВУ (${toMoscowList.size} изменений)", fontWeight = FontWeight.Bold, fontSize = 13.sp) }
                    )
                    Tab(
                        selected = (activeTab == 1),
                        onClick = { activeTab = 1 },
                        text = { Text("ОТ МОСКВЫ (${fromMoscowList.size} изменений)", fontWeight = FontWeight.Bold, fontSize = 13.sp) }
                    )
                }

                Spacer(modifier = Modifier.height(10.dp))

                // ТАБЛИЦА РЕЗУЛЬТАТОВ
                val currentDisplayList = if (activeTab == 0) toMoscowList else fromMoscowList
                Card(
                    modifier = Modifier.fillMaxSize(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = Theme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Column(modifier = Modifier.fillMaxSize().padding(10.dp)) {
                        Row(
                            modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF1F5F9)).padding(10.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text("ПОЕЗД", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = Theme.Navy, modifier = Modifier.width(110.dp))
                            Text("НАЗНАЧЕНИЕ", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = Theme.Navy, modifier = Modifier.width(160.dp))
                            Text("ВРЕМЯ АФИШИ", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = Theme.Navy, modifier = Modifier.width(120.dp))
                            Text("СТАТУС ИЗМЕНЕНИЯ", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = Theme.Navy, modifier = Modifier.weight(1f))
                        }

                        Divider(color = Theme.Border, modifier = Modifier.padding(vertical = 4.dp))

                        if (currentDisplayList.isEmpty()) {
                            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                Text("Нажмите «Сверить станцию» для формирования списка изменений.", color = Theme.TextMuted)
                            }
                        } else {
                            LazyColumn(modifier = Modifier.fillMaxSize()) {
                                items(currentDisplayList) { item ->
                                    val isCancel = item.isCancelled
                                    val bg = if (isCancel) Color(0xFFFFF5F5) else Color.White
                                    val border = if (isCancel) Color(0xFFFEE2E2) else Color(0xFFF1F5F9)

                                    Row(
                                        modifier = Modifier
                                            .fillMaxWidth()
                                            .padding(vertical = 3.dp)
                                            .clip(RoundedCornerShape(6.dp))
                                            .background(bg)
                                            .border(1.dp, border, RoundedCornerShape(6.dp))
                                            .padding(8.dp),
                                        verticalAlignment = Alignment.CenterVertically
                                    ) {
                                        Text(
                                            text = "№${item.trainNum}",
                                            fontWeight = FontWeight.Bold,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            color = Theme.Navy,
                                            modifier = Modifier.width(110.dp)
                                        )

                                        Text(
                                            text = item.destination,
                                            fontSize = 13.sp,
                                            color = Theme.TextDark,
                                            modifier = Modifier.width(160.dp)
                                        )

                                        Text(
                                            text = item.afishaTime,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            fontWeight = FontWeight.SemiBold,
                                            color = Theme.TextMuted,
                                            modifier = Modifier.width(120.dp)
                                        )

                                        Text(
                                            text = item.changeText,
                                            fontSize = 13.sp,
                                            fontWeight = if (isCancel) FontWeight.Bold else FontWeight.Medium,
                                            color = if (isCancel) Theme.Red else Theme.TextDark,
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
fun UploadCard(title: String, desc: String, isLoaded: Boolean, color: Color, icon: ImageVector, modifier: Modifier, onClick: () -> Unit) {
    Card(
        modifier = modifier.clickable { onClick() },
        shape = RoundedCornerShape(10.dp),
        colors = CardDefaults.cardColors(containerColor = Theme.CardBg),
        elevation = CardDefaults.cardElevation(1.dp)
    ) {
        Row(modifier = Modifier.fillMaxWidth().padding(12.dp), verticalAlignment = Alignment.CenterVertically) {
            Surface(shape = RoundedCornerShape(8.dp), color = if (isLoaded) color.copy(alpha = 0.12f) else Color(0xFFF1F5F9), modifier = Modifier.size(40.dp)) {
                Box(contentAlignment = Alignment.Center) {
                    Icon(if (isLoaded) Icons.Default.Check else icon, contentDescription = null, tint = if (isLoaded) color else Theme.TextMuted, modifier = Modifier.size(20.dp))
                }
            }
            Spacer(modifier = Modifier.width(10.dp))
            Column(modifier = Modifier.weight(1f)) {
                Text(title, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = Theme.TextMuted)
                Text(desc, fontSize = 12.sp, fontWeight = FontWeight.SemiBold, color = Theme.Navy, maxLines = 1)
            }
            TextButton(onClick = onClick) {
                Text(if (isLoaded) "Заменить" else "Выбрать", fontSize = 12.sp, fontWeight = FontWeight.Bold, color = color)
            }
        }
    }
}

// Полноценный анализ направления
fun processStationSchedule(
    afisha: File,
    telegram: File?,
    vygruzki: List<File>,
    targetStation: String
): Pair<List<TrainChangeRecord>, List<TrainChangeRecord>> {
    val toMsk = mutableListOf<TrainChangeRecord>()
    val fromMsk = mutableListOf<TrainChangeRecord>()

    // Читаем телеграмму
    val telText = buildString {
        telegram?.let { f ->
            try {
                if (f.name.endsWith(".docx", true)) {
                    FileInputStream(f).use { fis ->
                        XWPFDocument(fis).use { doc -> doc.paragraphs.forEach { append(it.text).append("\n") } }
                    }
                } else if (f.name.endsWith(".pdf", true)) {
                    Loader.loadPDF(f).use { doc -> append(PDFTextStripper().getText(doc)).append("\n") }
                }
            } catch (_: Exception) {}
        }
    }

    // Читаем выгрузки
    val vygText = buildString {
        vygruzki.forEach { f ->
            try { append(f.readText()).append("\n") } catch (_: Exception) {}
        }
    }

    // Читаем Excel Афиши
    try {
        FileInputStream(afisha).use { fis ->
            val wb = WorkbookFactory.create(fis)
            val timeFmt = SimpleDateFormat("HH:mm")

            for (sheetIndex in 0 until wb.numberOfSheets) {
                val sheet = wb.getSheetAt(sheetIndex)
                val sheetName = sheet.sheetName.lowercase()
                val isToMoscow = sheetName.contains("москва") && (sheetName.startsWith("тверь") || sheetName.contains("конаково"))
                val isFromMoscow = sheetName.startsWith("москва")

                if (!isToMoscow && !isFromMoscow) continue
                val dir = if (isToMoscow) Direction.TO_MOSCOW else Direction.FROM_MOSCOW

                // Ищем строку номеров поездов (обычно строка 3-4)
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

                // Ищем строку станции
                var stationRowIdx = -1
                for (r in trainRowIdx + 1..sheet.lastRowNum) {
                    val row = sheet.getRow(r) ?: continue
                    val name = row.getCell(0)?.toString()?.trim() ?: ""
                    if (name.contains(targetStation, ignoreCase = true) ||
                        (targetStation == "Зеленоград-Крюково" && (name.contains("Крюково", true) || name.contains("Зеленоград", true)))
                    ) {
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

                    val numOnly = rawNum.split("/").first()
                    var isCancelled = false
                    var newTime: String? = null
                    var changeText = ""

                    // 1. Поиск в выгрузках TXT
                    val vygRegex = Regex("№$numOnly.*?отправлением в [0-9:]+\\s*-\\s*(.*?)(?=\\n|\$)")
                    val vygMatch = vygRegex.find(vygText)

                    if (vygMatch != null) {
                        val desc = vygMatch.groupValues[1].trim()
                        if (desc.contains("ОТМЕНЕН", true)) {
                            isCancelled = true
                            changeText = "ОТМЕНЕН"
                        } else if (desc.contains("отправится в", true)) {
                            val tMatch = Regex("отправится в (\\d{2}:\\d{2})").find(desc)
                            newTime = tMatch?.groupValues?.get(1)
                            changeText = desc
                        } else {
                            changeText = desc
                        }
                    } else {
                        // 2. Поиск в телеграмме
                        if (telText.contains("П. НР $numOnly") || telText.contains("№ $numOnly") || telText.contains("№$numOnly")) {
                            val idx = telText.indexOf(numOnly)
                            val snippet = telText.substring((idx - 30).coerceAtLeast(0), (idx + 60).coerceAtMost(telText.length))
                            if (snippet.contains("отмен", true) || snippet.contains("ОТМЕН", true)) {
                                isCancelled = true
                                changeText = "ОТМЕНЕН"
                            }
                        }
                    }

                    // Если найдены изменения — добавляем в итоговый список
                    if (isCancelled || newTime != null || changeText.isNotEmpty()) {
                        val dest = if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Крюково/Твери"
                        val record = TrainChangeRecord(
                            trainNum = rawNum,
                            destination = dest,
                            afishaTime = afishaTime,
                            newTime = newTime,
                            isCancelled = isCancelled,
                            changeText = if (isCancelled) "ОТМЕНЕН" else changeText,
                            direction = dir
                        )
                        if (dir == Direction.TO_MOSCOW) toMsk.add(record) else fromMsk.add(record)
                    }
                }
            }
        }
    } catch (_: Exception) {}

    return Pair(toMsk, fromMsk)
}

// Экспорт двух файлов строго по шаблону "Молжаниново на Москву" и "Молжаниново от Москвы"
fun exportStationReports(
    station: String,
    toMoscow: List<TrainChangeRecord>,
    fromMoscow: List<TrainChangeRecord>
) {
    val desktop = File(System.getProperty("user.home"), "Desktop")

    // 1. Файл НА МОСКВУ
    val fileToMsk = File(desktop, "$station на Москву.txt")
    fileToMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station  Изменения в графике движения эл/поездов")
        out.println("на Ленинградский вокзал, Ипподром")
        out.println("15, 23 октября")
        toMoscow.forEach { r ->
            if (r.isCancelled) {
                out.println("        №${r.trainNum} ${r.destination} отправлением в ${r.afishaTime} - ОТМЕНЕН")
            } else {
                out.println("        №${r.trainNum} ${r.destination} отправлением в ${r.afishaTime} - ${r.changeText}")
            }
        }
    }

    // 2. Файл ОТ МОСКВЫ
    val fileFromMsk = File(desktop, "$station от Москвы.txt")
    fileFromMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station Изменения в графике движения эл/поездов")
        out.println("от Ленинградского вокзала, Ипподрома")
        out.println("15, 23 октября")
        fromMoscow.forEach { r ->
            if (r.isCancelled) {
                out.println("        №${r.trainNum} отправлением в ${r.afishaTime} - ОТМЕНЕН")
            } else {
                out.println("        №${r.trainNum} отправлением в ${r.afishaTime} - ${r.changeText}")
            }
        }
    }
}

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "ScheduleVerifier PRO — Ревизор расписаний с приоритетом телеграмм"
    ) {
        App()
    }
}
