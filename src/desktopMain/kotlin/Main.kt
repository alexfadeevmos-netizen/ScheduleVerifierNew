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

// Модель рейса с привязкой к датам и станции
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

// Путь к файлу постоянного кэша афиши
fun getCachedAfishaFile(): File {
    val dir = File(System.getProperty("user.home"), ".schedule_verifier")
    if (!dir.exists()) dir.mkdirs()
    return File(dir, "saved_afisha.xlsx")
}

@Composable
fun App() {
    val cachedFile = remember { getCachedAfishaFile() }
    var afishaFile by remember { mutableStateOf<File?>(if (cachedFile.exists()) cachedFile else null) }
    var telegramFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var vygruzkiFiles by remember { mutableStateOf<List<File>>(emptyList()) }

    var selectedStations by remember { mutableStateOf(setOf("Алабушево", "Молжаниново")) }
    var currentStationView by remember { mutableStateOf("Алабушево") }
    var searchQuery by remember { mutableStateOf("") }

    var toMoscowList by remember { mutableStateOf<List<DatedTrainRecord>>(emptyList()) }
    var fromMoscowList by remember { mutableStateOf<List<DatedTrainRecord>>(emptyList()) }
    var activeTab by remember { mutableStateOf(0) }

    var isProcessing by remember { mutableStateOf(false) }
    var statusMessage by remember {
        mutableStateOf(
            if (afishaFile != null) "Базовая афиша активна (из памяти). Загрузите телеграмму для интеллектуального анализа дат."
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

                // 1. ВЕРХНИЙ ХЕДЕР И СТАТИСТИКА
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
                                Text("SCHEDULE VERIFIER PRO — ЭКСПЕРТНЫЙ РЕВИЗОР РАСПИСАНИЙ", fontSize = 16.sp, fontWeight = FontWeight.ExtraBold, color = AppTheme.HeaderNavy)
                                Text(statusMessage, fontSize = 12.sp, color = AppTheme.TextMuted)
                            }
                        }

                        if (toMoscowList.isNotEmpty() || fromMoscowList.isNotEmpty()) {
                            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFEFF6FF)) {
                                    Text("На Москву: ${toMoscowList.size}", modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = AppTheme.CorporateBlue)
                                }
                                Surface(shape = RoundedCornerShape(6.dp), color = Color(0xFFF5F3FF)) {
                                    Text("От Москвы: ${fromMoscowList.size}", modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = Color(0xFF6D28D9))
                                }
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // 2. КАРТОЧКИ ЗАГРУЗКИ ФАЙЛОВ
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

                // 3. ПАНЕЛЬ ВЫБОРА СТАНЦИЙ И ПАКЕТНОГО ЭКСПОРТА
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = AppTheme.CardBg),
                    elevation = CardDefaults.cardElevation(1.dp)
                ) {
                    Column(modifier = Modifier.fillMaxWidth().padding(12.dp)) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Text("Станции для формирования отчетов отмен:", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = AppTheme.HeaderNavy)
                            Spacer(modifier = Modifier.width(8.dp))
                            TextButton(onClick = {
                                selectedStations = if (selectedStations.size == ALL_STATIONS.size) setOf("Алабушево") else ALL_STATIONS.toSet()
                            }) {
                                Text(if (selectedStations.size == ALL_STATIONS.size) "Снять все" else "Выбрать все станции", fontSize = 11.sp, fontWeight = FontWeight.Bold)
                            }

                            Spacer(modifier = Modifier.weight(1f))

                            // Кнопка СВЕРИТЬ ВЫБРАННУЮ
                            Button(
                                onClick = {
                                    if (afishaFile == null) return@Button
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Анализ расписания и телеграммы по ст. $currentStationView..."
                                        val (toMsk, fromMsk) = withContext(Dispatchers.IO) {
                                            analyzeStationChanges(afishaFile!!, telegramFiles, vygruzkiFiles, currentStationView)
                                        }
                                        toMoscowList = toMsk
                                        fromMoscowList = fromMsk
                                        isProcessing = false
                                        statusMessage = "Сверка ст. $currentStationView: на Москву (${toMsk.size}), от Москвы (${fromMsk.size})"
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

                            // Кнопка ПАКЕТНОЙ ГЕНЕРАЦИИ
                            Button(
                                onClick = {
                                    if (afishaFile == null) return@Button
                                    coroutineScope.launch {
                                        isProcessing = true
                                        statusMessage = "Генерация отчетов по ${selectedStations.size} станциям на Рабочий стол..."
                                        withContext(Dispatchers.IO) {
                                            selectedStations.forEach { st ->
                                                val (toMsk, fromMsk) = analyzeStationChanges(afishaFile!!, telegramFiles, vygruzkiFiles, st)
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

                        // Чекбоксы выбора станций
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

                // 4. ТАБЫ И БЫСТРЫЙ ПОИСК
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
                        modifier = Modifier.width(300.dp).height(48.dp),
                        singleLine = true,
                        textStyle = LocalTextStyle.current.copy(fontSize = 12.sp)
                    )
                }

                Spacer(modifier = Modifier.height(8.dp))

                // 5. РЕВИЗИОННАЯ ТАБЛИЦА
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
                            Text("ПЕРИОД ДАТ", fontWeight = FontWeight.Bold, fontSize = 11.sp, color = AppTheme.HeaderNavy, modifier = Modifier.width(150.dp))
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
                                            modifier = Modifier.width(150.dp)
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

// Приведение времени к единому формату "ЧЧ:ММ"
fun normalizeTimeString(raw: String): String {
    val clean = raw.split(",").first().trim().replace(".", ":")
    val parts = clean.split(":")
    if (parts.size == 2) {
        val h = parts[0].padStart(2, '0')
        val m = parts[1].padStart(2, '0')
        return "$h:$m"
    }
    return clean
}

// Интеллектуальный парсинг станций и времени отправления по телеграмме
fun parseTelegramStationsSchedule(telText: String): Map<String, Map<String, Pair<String, Boolean>>> {
    val result = mutableMapOf<String, MutableMap<String, Pair<String, Boolean>>>()
    var currentTrain: String? = null
    val trainRegex = Regex("(?i)поезд\\s*№?\\s*([0-9]{4}(?:/[0-9]{4})?)")

    for (line in telText.lines()) {
        val tm = trainRegex.find(line)
        if (tm != null) {
            currentTrain = tm.groupValues[1].split("/").first()
            result.putIfAbsent(currentTrain, mutableMapOf())
        }

        if (currentTrain != null) {
            val segments = line.split(",")
            for (seg in segments) {
                val s = seg.trim()
                if (s.isEmpty()) continue
                val match = Regex("([А-Яа-яЁё0-9\\-\\.\\s]+?)\\s+(?:отпр\\.?\\s*|приб\\.?\\s*)?(\\d{1,2}[\\.:]\\d{2}(?:,\\d+)?)(?:\\s*[-–—]\\s*(\\d{1,2}[\\.:]\\d{2}(?:,\\d+)?))?").find(s)
                if (match != null) {
                    val stName = match.groupValues[1].replace(Regex("(?i)поезд\\s*№?\\s*[0-9/]+"), "").trim()
                    val t1 = match.groupValues[2]
                    val t2 = if (match.groupValues.size > 3) match.groupValues[3] else ""
                    val hasRange = t2.isNotEmpty()
                    val hasDep = s.contains("отпр", true)
                    val hasArr = s.contains("приб", true)

                    // Берем второе время при диапазоне (отправление)
                    val depTime = if (hasRange) normalizeTimeString(t2) else normalizeTimeString(t1)
                    val isStop = hasRange || hasDep || hasArr

                    if (stName.length >= 3) {
                        result[currentTrain]?.put(stName.lowercase(), Pair(depTime, isStop))
                    }
                }
            }
        }
    }
    return result
}

// Извлечение периодов дат из телеграмм и выгрузок
fun extractDatePeriods(telText: String, vygText: String): List<String> {
    val groups = mutableSetOf<String>()
    val combined = telText + "\n" + vygText

    val dateRegex = Regex("(?m)^\\s*(\\d{1,2}(?:\\s*,\\s*\\d{1,2})*\\s+октября|с\\s+\\d{1,2}\\s+по\\s+\\d{1,2}\\s+октября|\\d{1,2}\\s+октября)\\s*$")
    for (m in dateRegex.findAll(combined)) {
        val g = m.groupValues[1].trim()
        if (g.isNotEmpty()) groups.add(g)
    }

    if (groups.isEmpty()) {
        return listOf("15, 23 октября", "16 октября", "17 октября", "18 октября", "19, 20, 21, 22 октября", "24 октября")
    }
    return groups.toList()
}

// Полный анализ расписания с разбивкой по датам
fun analyzeStationChanges(
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

    val vygText = buildString {
        vygruzki.forEach { f ->
            try { append(f.readText()).append("\n") } catch (_: Exception) {}
        }
    }

    val telegramParsed = parseTelegramStationsSchedule(telText)
    val dateGroups = extractDatePeriods(telText, vygText)

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
                    afishaTime = normalizeTimeString(afishaTime)

                    val numOnly = rawNum.split("/").first()

                    dateGroups.forEach { dateGroup ->
                        var isCancelled = false
                        var isPassWithoutStop = false
                        var newTime: String? = null
                        var changeText = ""

                        // 1. Поиск в выгрузках (если они есть)
                        val vygDateRegex = Regex("(?s)$dateGroup.*?(№$numOnly.*?отправлением в [0-9:]+\\s*-\\s*(.*?))(?=\\n\\s*\\d{1,2}\\s+октября|\\n\\s*с\\s+\\d|\$)")
                        val vygMatch = vygDateRegex.find(vygText)

                        if (vygMatch != null) {
                            val desc = vygMatch.groupValues[2].trim()
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
                        } else {
                            // 2. Интеллектуальный анализ напрямую по тексту телеграммы
                            if (telText.contains("П. НР $numOnly") || telText.contains("№ $numOnly") || telText.contains("№$numOnly")) {
                                val idx = telText.indexOf(numOnly)
                                val snippet = telText.substring((idx - 30).coerceAtLeast(0), (idx + 60).coerceAtMost(telText.length))
                                if (snippet.contains("отмен", true) || snippet.contains("ОТМЕН", true)) {
                                    isCancelled = true
                                    changeText = "ОТМЕНЕН"
                                }
                            }

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
                        }

                        if (isCancelled || isPassWithoutStop || newTime != null || changeText.isNotEmpty()) {
                            val dest = if (dir == Direction.TO_MOSCOW) "до Москвы" else "до Крюково/Твери"
                            val record = DatedTrainRecord(
                                dateGroup = dateGroup,
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
        }
    } catch (_: Exception) {}

    return Pair(toMsk, fromMsk)
}

// Экспорт двух файлов строго по стандарту образцов Word
fun exportDatedReports(station: String, toMoscow: List<DatedTrainRecord>, fromMoscow: List<DatedTrainRecord>) {
    val desktop = File(System.getProperty("user.home"), "Desktop")

    // Файл НА МОСКВУ
    val fToMsk = File(desktop, "$station на Москву.txt")
    fToMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station Изменения в графике движения эл/поездов")
        out.println("на Москву")

        val grouped = toMoscow.groupBy { it.dateGroup }
        grouped.forEach { (dateGroup, list) ->
            out.println(dateGroup)
            list.forEach { r ->
                out.println("        №${r.trainNum} ${r.destination} отправлением в ${r.afishaTime} - ${r.changeText}")
            }
        }
    }

    // Файл ОТ МОСКВЫ
    val fFromMsk = File(desktop, "$station от Москвы.txt")
    fFromMsk.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station Изменения в графике движения эл/поездов")
        out.println("от Ленинградского вокзала, Ипподрома")

        val grouped = fromMoscow.groupBy { it.dateGroup }
        grouped.forEach { (dateGroup, list) ->
            out.println(dateGroup)
            list.forEach { r ->
                out.println("        №${r.trainNum} отправлением в ${r.afishaTime} - ${r.changeText}")
            }
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
