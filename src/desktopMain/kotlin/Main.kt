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

data class TrainChangeRecord(
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

object Theme {
    val Bg = Color(0xFFF8FAFC)
    val CardBg = Color(0xFFFFFFFF)
    val Navy = Color(0xFF0F172A)
    val PrimaryBlue = Color(0xFF1E3A8A)
    val Green = Color(0xFF059669)
    val Red = Color(0xFFDC2626)
    val Border = Color(0xFFE2E8F0)
    val TextDark = Color(0xFF1E293B)
    val TextMuted = Color(0xFF64748B)
}

val ALL_STATIONS = listOf(
    "Молжаниново", "Зеленоград-Крюково", "Химки", "Сходня", "Фирсановская",
    "Подрезково", "Новоподрезково", "Левобережная", "Ховрино", "Грачёвская",
    "Моссельмаш", "Лихоборы", "Останкино", "Подсолнечная", "Клин", "Тверь"
)

// Папка для постоянного хранения афиши
fun getCachedAfishaFile(): File {
    val appDir = File(System.getProperty("user.home"), ".schedule_verifier")
    if (!appDir.exists()) appDir.mkdirs()
    return File(appDir, "saved_afisha.xlsx")
}

@Composable
fun App() {
    val cachedFile = remember { getCachedAfishaFile() }
    var afishaFile by remember { mutableStateOf<File?>(if (cachedFile.exists()) cachedFile else null) }
    var telegramFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var vygruzkiFiles by remember { mutableStateOf<List<File>>(emptyList()) }

    var selectedStations by remember { mutableStateOf(setOf("Молжаниново")) }
    var currentStationView by remember { mutableStateOf("Молжаниново") }
    var searchQuery by remember { mutableStateOf("") }

    var toMoscowList by remember { mutableStateOf<List<TrainChangeRecord>>(emptyList()) }
    var fromMoscowList by remember { mutableStateOf<List<TrainChangeRecord>>(emptyList()) }
    var activeTab by remember { mutableStateOf(0) }

    var isProcessing by remember { mutableStateOf(false) }
    var statusMessage by remember {
        mutableStateOf(
            if (afishaFile != null) "Базовая афиша активна (из кэша). Загрузите телеграммы для сверки."
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
        Surface(modifier = Modifier.fillMaxSize(), color = Theme.Bg) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {

                // ШАПКА ПРИЛОЖЕНИЯ
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = Theme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(14.dp),
                        verticalAlignment = Alignment.CenterVertically,
                        horizontalArrangement = Arrangement.SpaceBetween
                    ) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Icon(Icons.Default.CheckCircle, contentDescription = null, tint = Theme.PrimaryBlue, modifier = Modifier.size(26.dp))
                            Spacer(modifier = Modifier.width(10.dp))
                            Column {
                                Text("SCHEDULE VERIFIER PRO — РЕВИЗОР И ГЕНЕРАТОР ОТМЕН", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = Theme.Navy)
                                Text(statusMessage, fontSize = 12.sp, color = Theme.TextMuted)
                            }
                        }

                        if (toMoscowList.isNotEmpty() || fromMoscowList.isNotEmpty()) {
                            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFEFF6FF)) {
                                    Text("На Москву: ${toMoscowList.size}", modifier = Modifier.padding(8.dp, 4.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = Theme.PrimaryBlue)
                                }
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFF5F3FF)) {
                                    Text("От Москвы: ${fromMoscowList.size}", modifier = Modifier.padding(8.dp, 4.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = Color(0xFF6D28D9))
                                }
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // КАРТОЧКИ ЗАГРУЗКИ ФАЙЛОВ
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                    UploadCard(
                        title = "1. АФИША (ПОСТОЯННАЯ)",
                        desc = if (afishaFile != null) "✓ Активна: ${afishaFile!!.name}" else "Выберите файл Excel (.xlsx)",
                        isLoaded = afishaFile != null,
                        color = Theme.Green,
                        icon = Icons.Default.DateRange,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Афишу (.xlsx)")
                        if (f.isNotEmpty()) {
                            val selected = f.first()
                            // Кэшируем афишу на постоянное хранение
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

                    UploadCard(
                        title = "2. ТЕЛЕГРАММЫ",
                        desc = if (telegramFiles.isNotEmpty()) "Выбрано документов: ${telegramFiles.size}" else "Word (.docx) / PDF",
                        isLoaded = telegramFiles.isNotEmpty(),
                        color = Theme.PrimaryBlue,
                        icon = Icons.Default.Info,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Телеграммы", multiple = true)
                        if (f.isNotEmpty()) {
                            telegramFiles = f
                            statusMessage = "Загружено телеграмм: ${f.size}"
                        }
                    }

                    UploadCard(
                        title = "3. ОПЕРАТИВНЫЕ ВЫГРУЗКИ",
                        desc = if (vygruzkiFiles.isNotEmpty()) "Выбрано выгрузок: ${vygruzkiFiles.size}" else "Текстовые файлы (.txt)",
                        isLoaded = vygruzkiFiles.isNotEmpty(),
                        color = Color(0xFF6366F1),
                        icon = Icons.Default.List,
                        modifier = Modifier.weight(1f)
                    ) {
                        val f = pickFiles("Выберите Выгрузки (.txt)", multiple = true)
                        if (f.isNotEmpty()) {
                            vygruzkiFiles = f
                            statusMessage = "Загружено оперативных выгрузок: ${f.size}"
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // ПАНЕЛЬ СТАНЦИЙ И ДЕЙСТВИЙ
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = Theme.CardBg),
                    elevation = CardDefaults.cardElevation(1.dp)
                ) {
                    Column(modifier = Modifier.fillMaxWidth().padding(12.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text("Станции для формирования файлов отмен:", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = Theme.Navy)
                            Spacer(modifier = Modifier.width(8.dp))
                            TextButton(onClick = {
                                selectedStations = if (selectedStations.size == ALL_STATIONS.size) setOf("Молжаниново") else ALL_STATIONS.toSet()
                            }) {
                                Text(if (selectedStations.size == ALL_STATIONS.size) "Снять все" else "Выбрать все станции", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }

                            Spacer(modifier = Modifier.weight(1f))

                            // Кнопка СВЕРИТЬ
                            Button(
                                onClick = {
                                    if (afishaFile == null) return@Button
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Идет интеллектуальный анализ расписания..."
                                        val (toMsk, fromMsk) = withContext(Dispatchers.IO) {
                                            processStation(afishaFile!!, telegramFiles, vygruzkiFiles, currentStationView)
                                        }
                                        toMoscowList = toMsk
                                        fromMoscowList = fromMsk
                                        isProcessing = false
                                        statusMessage = "Сверка ст. $currentStationView завершена: на Москву (${toMsk.size}), от Москвы (${fromMsk.size})"
                                    }
                                },
                                enabled = !isProcessing && afishaFile != null,
                                colors = ButtonDefaults.buttonColors(containerColor = Theme.PrimaryBlue),
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text(if (isProcessing) "Сверка..." else "Сверить $currentStationView")
                            }

                            Spacer(modifier = Modifier.width(8.dp))

                            // Кнопка ПАКЕТНОЙ ГЕНЕРАЦИИ
                            Button(
                                onClick = {
                                    if (afishaFile == null) return@Button
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Пакетная генерация файлов отмен на Рабочий стол..."
                                        withContext(Dispatchers.IO) {
                                            selectedStations.forEach { st ->
                                                val (toMsk, fromMsk) = processStation(afishaFile!!, telegramFiles, vygruzkiFiles, st)
                                                exportStationReports(st, toMsk, fromMsk)
                                            }
                                        }
                                        isProcessing = false
                                        statusMessage = "УСПЕШНО! Сгенерированы файлы отмен по ${selectedStations.size} станциям на Рабочем столе!"
                                    }
                                },
                                enabled = !isProcessing && afishaFile != null && selectedStations.isNotEmpty(),
                                colors = ButtonDefaults.buttonColors(containerColor = Theme.Red),
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.Share, contentDescription = null, modifier = Modifier.size(16.dp))
                                Spacer(modifier = Modifier.width(6.dp))
                                Text("Сгенерировать файлы ОТМЕНА (${selectedStations.size} ст.)")
                            }
                        }

                        Spacer(modifier = Modifier.height(6.dp))

                        // Чекбоксы станций
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                            ALL_STATIONS.take(8).forEach { st ->
                                StationChip(st, selectedStations.contains(st), currentStationView == st,
                                    onToggle = {
                                        selectedStations = if (selectedStations.contains(st)) selectedStations - st else selectedStations + st
                                    },
                                    onSelectView = {
                                        currentStationView = st
                                    }
                                )
                            }
                        }
                        Spacer(modifier = Modifier.height(4.dp))
                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(4.dp)) {
                            ALL_STATIONS.drop(8).forEach { st ->
                                StationChip(st, selectedStations.contains(st), currentStationView == st,
                                    onToggle = {
                                        selectedStations = if (selectedStations.contains(st)) selectedStations - st else selectedStations + st
                                    },
                                    onSelectView = {
                                        currentStationView = st
                                    }
                                )
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // ТАБЫ И ПОИСК
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween, verticalAlignment = Alignment.CenterVertically) {
                    TabRow(
                        selectedTabIndex = activeTab,
                        containerColor = Theme.CardBg,
                        contentColor = Theme.PrimaryBlue,
                        modifier = Modifier.width(420.dp).clip(RoundedCornerShape(8.dp))
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
                        placeholder = { Text("Поиск нитки поезда (напр. 6501)...", fontSize = 12.sp) },
                        modifier = Modifier.width(300.dp).height(48.dp),
                        singleLine = true,
                        textStyle = LocalTextStyle.current.copy(fontSize = 12.sp)
                    )
                }

                Spacer(modifier = Modifier.height(8.dp))

                // СПИСОК РЕЗУЛЬТАТОВ
                val currentList = (if (activeTab == 0) toMoscowList else fromMoscowList)
                    .filter { it.trainNum.contains(searchQuery.trim(), ignoreCase = true) }

                Card(
                    modifier = Modifier.fillMaxSize(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = Theme.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Column(modifier = Modifier.fillMaxSize().padding(10.dp)) {
                        Row(
                            modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(6.dp)).background(Color(0xFFF1F5F9)).padding(8.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text("ПОЕЗД", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = Theme.Navy, modifier = Modifier.width(100.dp))
                            Text("НАЗНАЧЕНИЕ", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = Theme.Navy, modifier = Modifier.width(160.dp))
                            Text("ПО АФИШЕ", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = Theme.Navy, modifier = Modifier.width(110.dp))
                            Text("ИЗМЕНЕНИЯ ПО ТЕЛЕГРАММЕ / ВЫГРУЗКЕ", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = Theme.Navy, modifier = Modifier.weight(1f))
                        }

                        Divider(color = Theme.Border, modifier = Modifier.padding(vertical = 4.dp))

                        if (currentList.isEmpty()) {
                            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                Text("Нет данных для отображения по ст. $currentStationView. Нажмите «Сверить».", color = Theme.TextMuted, fontSize = 13.sp)
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
                                            text = "№${item.trainNum}",
                                            fontWeight = FontWeight.Bold,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            color = Theme.Navy,
                                            modifier = Modifier.width(100.dp)
                                        )

                                        Text(
                                            text = item.destination,
                                            fontSize = 12.sp,
                                            color = Theme.TextDark,
                                            modifier = Modifier.width(160.dp)
                                        )

                                        Text(
                                            text = item.afishaTime,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            fontWeight = FontWeight.SemiBold,
                                            color = Theme.TextMuted,
                                            modifier = Modifier.width(110.dp)
                                        )

                                        Text(
                                            text = item.changeText,
                                            fontSize = 12.sp,
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
fun StationChip(name: String, isChecked: Boolean, isCurrentView: Boolean, onToggle: () -> Unit, onSelectView: () -> Unit) {
    Surface(
        shape = RoundedCornerShape(6.dp),
        color = if (isCurrentView) Theme.PrimaryBlue.copy(alpha = 0.12f) else Color(0xFFF8FAFC),
        border = androidx.compose.foundation.BorderStroke(1.dp, if (isCurrentView) Theme.PrimaryBlue else Theme.Border),
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
                color = if (isCurrentView) Theme.PrimaryBlue else Theme.TextDark
            )
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
            Surface(shape = RoundedCornerShape(8.dp), color = if (isLoaded) color.copy(alpha = 0.12f) else Color(0xFFF1F5F9), modifier = Modifier.size(38.dp)) {
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
                Text(if (isLoaded) "Заменить" else "Выбрать", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = color)
            }
        }
    }
}

// Нормализатор времени: "12.49,5" -> "12:49"
fun normalizeTime(raw: String): String {
    val clean = raw.split(",").first().trim().replace(".", ":")
    val parts = clean.split(":")
    if (parts.size == 2) {
        val h = parts[0].padStart(2, '0')
        val m = parts[1].padStart(2, '0')
        return "$h:$m"
    }
    return clean
}

// Интеллектуальный разбор телеграммы по поездам и станциям
fun parseTelegramSchedule(telText: String): Map<String, Map<String, Pair<String, Boolean>>> {
    // trainNum -> (stationName -> Pair(departureTime, isStop))
    val result = mutableMapOf<String, MutableMap<String, Pair<String, Boolean>>>()

    val lines = telText.lines()
    var currentTrain: String? = null

    val trainHeaderRegex = Regex("(?i)поезд\\s*№?\\s*([0-9]{4}(?:/[0-9]{4})?)")

    for (line in lines) {
        val trainMatch = trainHeaderRegex.find(line)
        if (trainMatch != null) {
            currentTrain = trainMatch.groupValues[1].split("/").first()
            result.putIfAbsent(currentTrain, mutableMapOf())
        }

        if (currentTrain != null) {
            // Разделяем элементы маршрута (по запятым)
            val segments = line.split(",")
            for (seg in segments) {
                val s = seg.trim()
                if (s.isEmpty()) continue

                // Паттерн: Название_станции [отпр./приб.] время1 [- время2]
                val match = Regex("([А-Яа-яЁё0-9\\-\\.\\s]+?)\\s+(?:отпр\\.?\\s*|приб\\.?\\s*)?(\\d{1,2}[\\.:]\\d{2}(?:,\\d+)?)(?:\\s*[-–—]\\s*(\\d{1,2}[\\.:]\\d{2}(?:,\\d+)?))?").find(s)
                if (match != null) {
                    val stName = match.groupValues[1]
                        .replace(Regex("(?i)поезд\\s*№?\\s*[0-9/]+"), "")
                        .replace("отпр.", "")
                        .replace("приб.", "")
                        .trim()

                    val t1 = match.groupValues[2]
                    val t2 = if (match.groupValues.size > 3) match.groupValues[3] else ""

                    val hasRange = t2.isNotEmpty()
                    val hasDepWord = s.contains("отпр", ignoreCase = true)
                    val hasArrWord = s.contains("приб", ignoreCase = true)

                    // Правило: если диапазон (12:49-12:50), берем строго второе время (отправление)
                    val depTime = if (hasRange) normalizeTime(t2) else normalizeTime(t1)

                    // Одно время без диапазона и без "отпр" означает безостановочный проход
                    val isStop = hasRange || hasDepWord || hasArrWord

                    if (stName.length >= 3) {
                        result[currentTrain]?.put(stName.lowercase(), Pair(depTime, isStop))
                    }
                }
            }
        }
    }
    return result
}

// Анализ станции
fun processStation(
    afisha: File,
    telegrams: List<File>,
    vygruzki: List<File>,
    targetStation: String
): Pair<List<TrainChangeRecord>, List<TrainChangeRecord>> {
    val toMsk = mutableListOf<TrainChangeRecord>()
    val fromMsk = mutableListOf<TrainChangeRecord>()

    // Читаем все телеграммы
    val telText = buildString {
        telegrams.forEach { f ->
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

    // Парсим расписания телеграммы
    val telegramParsed = parseTelegramSchedule(telText)

    // Читаем выгрузки
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
                val sheetName = sheet.sheetName.lowercase()
                val isToMoscow = sheetName.contains("москва") && (sheetName.startsWith("тверь") || sheetName.contains("конаково"))
                val isFromMoscow = sheetName.startsWith("москва")

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
                    afishaTime = normalizeTime(afishaTime)

                    val numOnly = rawNum.split("/").first()
                    var isCancelled = false
                    var isPassWithoutStop = false
                    var newTime: String? = null
                    var changeText = ""

                    // 1. Проверка отмены в блоке "Отменить" телеграммы
                    if (telText.contains("П. НР $numOnly") || telText.contains("№ $numOnly") || telText.contains("№$numOnly")) {
                        val idx = telText.indexOf(numOnly)
                        val snippet = telText.substring((idx - 30).coerceAtLeast(0), (idx + 60).coerceAtMost(telText.length))
                        if (snippet.contains("отмен", true) || snippet.contains("ОТМЕН", true)) {
                            isCancelled = true
                            changeText = "ОТМЕНЕН"
                        }
                    }

                    // 2. Проверка в расписании телеграммы (по станциям)
                    if (!isCancelled && telegramParsed.containsKey(numOnly)) {
                        val stMap = telegramParsed[numOnly]!!
                        val key = stMap.keys.firstOrNull { it.contains(targetStation.lowercase()) || (targetStation == "Зеленоград-Крюково" && it.contains("крюково")) }
                        if (key != null) {
                            val (telDepTime, isStop) = stMap[key]!!
                            if (!isStop) {
                                isPassWithoutStop = true
                                changeText = "ОТМЕНЕН (проследует без остановки)"
                            } else if (telDepTime != afishaTime && telDepTime.isNotEmpty()) {
                                newTime = telDepTime
                                changeText = "отправится в $telDepTime"
                            }
                        }
                    }

                    // 3. Проверка в оперативных выгрузках TXT
                    val vygRegex = Regex("№$numOnly.*?отправлением в [0-9:]+\\s*-\\s*(.*?)(?=\\n|\$)")
                    val vygMatch = vygRegex.find(vygText)
                    if (vygMatch != null) {
                        val desc = vygMatch.groupValues[1].trim()
                        if (desc.contains("ОТМЕНЕН", true)) {
                            if (desc.contains("без остановки", true)) {
                                isPassWithoutStop = true
                                changeText = "ОТМЕНЕН (проследует без остановки)"
                            } else {
                                isCancelled = true
                                changeText = "ОТМЕНЕН"
                            }
                        } else if (desc.contains("отправится в", true)) {
                            val tMatch = Regex("отправится в (\\d{2}:\\d{2})").find(desc)
                            newTime = tMatch?.groupValues?.get(1)
                            changeText = desc
                        } else {
                            changeText = desc
                        }
                    }

                    if (isCancelled || isPassWithoutStop || newTime != null || changeText.isNotEmpty()) {
                        val dest = if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Крюково/Твери"
                        val record = TrainChangeRecord(
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
            }
        }
    } catch (_: Exception) {}

    return Pair(toMsk, fromMsk)
}

// Экспорт двух файлов строго по стандарту
fun exportStationReports(station: String, toMoscow: List<TrainChangeRecord>, fromMoscow: List<TrainChangeRecord>) {
    val desktop = File(System.getProperty("user.home"), "Desktop")

    // Файл НА МОСКВУ
    val fToMsk = File(desktop, "$station на Москву.txt")
    fToMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station  Изменения в графике движения эл/поездов")
        out.println("на Ленинградский вокзал, Ипподром")
        out.println("15, 23 октября")
        toMoscow.forEach { r ->
            out.println("        №${r.trainNum} ${r.destination} отправлением в ${r.afishaTime} - ${r.changeText}")
        }
    }

    // Файл ОТ МОСКВЫ
    val fFromMsk = File(desktop, "$station от Москвы.txt")
    fFromMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station Изменения в графике движения эл/поездов")
        out.println("от Ленинградского вокзала, Ипподрома")
        out.println("15, 23 октября")
        fromMoscow.forEach { r ->
            out.println("        №${r.trainNum} отправлением в ${r.afishaTime} - ${r.changeText}")
        }
    }
}

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "ScheduleVerifier PRO — Экспертный ревизор расписаний"
    ) {
        App()
    }
}
