import androidx.compose.foundation.background
import androidx.compose.foundation.border
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

// Модель элемента расхождения
data class DiffItem(
    val trainNum: String,
    val station: String,
    val afishaTime: String,
    val actualTime: String,
    val status: String,
    val details: String,
    val statusType: StatusType
)

enum class StatusType {
    NORMAL, MODIFIED, CANCELLED
}

// Корпоративная палитра железнодорожного ревизора
object ThemeColors {
    val BgCanvas = Color(0xFFF8FAFC)       // Мягкий фон Slate-50
    val CardBg = Color(0xFFFFFFFF)         // Чистый белый
    val HeaderSlate = Color(0xFF0F172A)    // Глубокий темно-синий
    val BorderSubtle = Color(0xFFE2E8F0)   // Границы Slate-200
    val TextMain = Color(0xFF1E293B)       // Основной текст
    val TextMuted = Color(0xFF64748B)      // Второстепенный текст

    val PrimaryNavy = Color(0xFF1E3A8A)    // Синий индиго
    val EmeraldGreen = Color(0xFF059669)   // Зеленый подтверждения
    val DangerRed = Color(0xFFDC2626)      // Красный отмены
    val AmberWarning = Color(0xFFD97706)   // Янтарный изменения
}

@Composable
fun App() {
    var afishaFile by remember { mutableStateOf<File?>(null) }
    var telegramFile by remember { mutableStateOf<File?>(null) }
    var vygruzkiFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var diffList by remember { mutableStateOf<List<DiffItem>>(emptyList()) }
    var selectedStation by remember { mutableStateOf("Зеленоград-Крюково") }
    var isProcessing by remember { mutableStateOf(false) }
    var statusLog by remember { mutableStateOf("Готов к работе. Загрузите файлы афиши, телеграммы и выгрузок.") }

    val coroutineScope = rememberCoroutineScope()
    val stationsList = listOf("Зеленоград-Крюково", "Химки", "Сходня", "Тверь", "Клин", "Подсолнечная", "Москва Пассажирская")

    fun pickFile(title: String, multiple: Boolean = false): List<File> {
        val d = FileDialog(null as Frame?, title, FileDialog.LOAD).apply {
            isMultipleMode = multiple
            isVisible = true
        }
        val res = d.files.toList()
        d.dispose()
        return res
    }

    val cancelledCount = diffList.count { it.statusType == StatusType.CANCELLED }
    val modifiedCount = diffList.count { it.statusType == StatusType.MODIFIED }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = ThemeColors.BgCanvas) {
            Column(modifier = Modifier.fillMaxSize().padding(18.dp)) {

                // 1. ВЕРХНИЙ ХЕДЕР: Название и статус-панель
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = ThemeColors.CardBg),
                    elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(16.dp),
                        verticalAlignment = Alignment.CenterVertically,
                        horizontalArrangement = Arrangement.SpaceBetween
                    ) {
                        Column {
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Icon(
                                    imageVector = Icons.Default.CheckCircle,
                                    contentDescription = null,
                                    tint = ThemeColors.PrimaryNavy,
                                    modifier = Modifier.size(26.dp)
                                )
                                Spacer(modifier = Modifier.width(10.dp))
                                Text(
                                    text = "SCHEDULE VERIFIER PRO",
                                    fontSize = 18.sp,
                                    fontWeight = FontWeight.ExtraBold,
                                    color = ThemeColors.HeaderSlate,
                                    letterSpacing = 1.sp
                                )
                                Spacer(modifier = Modifier.width(8.dp))
                                Surface(
                                    shape = RoundedCornerShape(4.dp),
                                    color = Color(0xFFEEF2F6)
                                ) {
                                    Text(
                                        text = "РЕВИЗОР РАСПИСАНИЙ",
                                        fontSize = 11.sp,
                                        fontWeight = FontWeight.Bold,
                                        color = ThemeColors.TextMuted,
                                        modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp)
                                    )
                                }
                            }
                            Spacer(modifier = Modifier.height(4.dp))
                            Text(
                                text = statusLog,
                                fontSize = 13.sp,
                                color = if (isProcessing) ThemeColors.PrimaryNavy else ThemeColors.TextMuted
                            )
                        }

                        // Мини-виджеты статистики
                        if (diffList.isNotEmpty()) {
                            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                                StatBadge("ОТМЕНЫ", cancelledCount.toString(), ThemeColors.DangerRed, Color(0xFFFEE2E2))
                                StatBadge("СДВИГИ", modifiedCount.toString(), ThemeColors.AmberWarning, Color(0xFFFEF3C7))
                            }
                        }
                    }
                }

                Spacer(modifier = Modifier.height(14.dp))

                // 2. БЛОК ЗАГРУЗКИ ДОКУМЕНТОВ
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    ActionUploadCard(
                        title = "АФИША РАСПИСАНИЯ",
                        subtitle = if (afishaFile != null) afishaFile!!.name else "Файл Excel (.xlsx)",
                        icon = Icons.Default.DateRange,
                        isLoaded = afishaFile != null,
                        color = ThemeColors.EmeraldGreen,
                        modifier = Modifier.weight(1f),
                        onClick = {
                            val f = pickFile("Выберите Афишу расписания (.xlsx)")
                            if (f.isNotEmpty()) {
                                afishaFile = f.first()
                                statusLog = "Афиша загружена: ${afishaFile?.name}"
                            }
                        }
                    )

                    ActionUploadCard(
                        title = "ОФИЦИАЛЬНАЯ ТЕЛЕГРАММА",
                        subtitle = if (telegramFile != null) telegramFile!!.name else "Word (.docx) или PDF",
                        icon = Icons.Default.Info,
                        isLoaded = telegramFile != null,
                        color = ThemeColors.PrimaryNavy,
                        modifier = Modifier.weight(1f),
                        onClick = {
                            val f = pickFile("Выберите Телеграмму (Word/PDF)")
                            if (f.isNotEmpty()) {
                                telegramFile = f.first()
                                statusLog = "Телеграмма загружена: ${telegramFile?.name}"
                            }
                        }
                    )

                    ActionUploadCard(
                        title = "ВЫГРУЗКИ ПО СТАНЦИЯМ",
                        subtitle = if (vygruzkiFiles.isNotEmpty()) "Выбрано файлов: ${vygruzkiFiles.size}" else "Текстовые файлы (.txt)",
                        icon = Icons.Default.List,
                        isLoaded = vygruzkiFiles.isNotEmpty(),
                        color = Color(0xFF4F46E5),
                        modifier = Modifier.weight(1f),
                        onClick = {
                            val f = pickFile("Выберите текстовые выгрузки (.txt)", multiple = true)
                            if (f.isNotEmpty()) {
                                vygruzkiFiles = f
                                statusLog = "Загружено оперативных выгрузок: ${f.size}"
                            }
                        }
                    )
                }

                Spacer(modifier = Modifier.height(14.dp))

                // 3. ПАНЕЛЬ ФИЛЬТРА СТАНЦИЙ И КНОПКИ ДЕЙСТВИЯ
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(10.dp),
                    colors = CardDefaults.cardColors(containerColor = ThemeColors.CardBg),
                    elevation = CardDefaults.cardElevation(1.dp)
                ) {
                    Row(
                        modifier = Modifier.fillMaxWidth().padding(horizontal = 14.dp, vertical = 10.dp),
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Text(
                            text = "Станция сверки:",
                            fontWeight = FontWeight.Bold,
                            fontSize = 13.sp,
                            color = ThemeColors.HeaderSlate
                        )
                        Spacer(modifier = Modifier.width(10.dp))

                        // Чипы станций
                        Row(horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                            stationsList.forEach { st ->
                                val isSelected = (selectedStation == st)
                                FilterChip(
                                    selected = isSelected,
                                    onClick = { selectedStation = st },
                                    label = {
                                        Text(
                                            text = st.replace("Зеленоград-", ""),
                                            fontSize = 12.sp,
                                            fontWeight = if (isSelected) FontWeight.Bold else FontWeight.Medium
                                        )
                                    },
                                    colors = FilterChipDefaults.filterChipColors(
                                        selectedContainerColor = ThemeColors.HeaderSlate,
                                        selectedLabelColor = Color.White
                                    )
                                )
                            }
                        }

                        Spacer(modifier = Modifier.weight(1f))

                        // Кнопка СВЕРИТЬ
                        Button(
                            onClick = {
                                if (afishaFile == null) return@Button
                                coroutineScope.launch {
                                    isProcessing = true
                                    statusLog = "Синхронизация Афиши, Телеграммы и Выгрузок..."
                                    diffList = withContext(Dispatchers.IO) {
                                        compareAllSources(afishaFile!!, telegramFile, vygruzkiFiles, selectedStation)
                                    }
                                    isProcessing = false
                                    statusLog = "Сверка завершена. Найдено расхождений по ст. $selectedStation: ${diffList.size}"
                                }
                            },
                            enabled = !isProcessing && afishaFile != null,
                            shape = RoundedCornerShape(8.dp),
                            colors = ButtonDefaults.buttonColors(containerColor = ThemeColors.PrimaryNavy),
                            contentPadding = PaddingValues(horizontal = 16.dp, vertical = 10.dp)
                        ) {
                            Icon(Icons.Default.Refresh, contentDescription = null, modifier = Modifier.size(16.dp))
                            Spacer(modifier = Modifier.width(6.dp))
                            Text(if (isProcessing) "Сверка..." else "Сверить данные", fontWeight = FontWeight.SemiBold)
                        }

                        Spacer(modifier = Modifier.width(8.dp))

                        // Кнопка ЭКСПОРТ В ФАЙЛ
                        Button(
                            onClick = {
                                if (diffList.isEmpty()) return@Button
                                exportCancellationReport(diffList, selectedStation)
                                statusLog = "Отчет 'ОТМЕНА_$selectedStation.txt' успешно сохранен на Рабочий стол!"
                            },
                            enabled = diffList.isNotEmpty(),
                            shape = RoundedCornerShape(8.dp),
                            colors = ButtonDefaults.buttonColors(containerColor = ThemeColors.DangerRed),
                            contentPadding = PaddingValues(horizontal = 16.dp, vertical = 10.dp)
                        ) {
                            Icon(Icons.Default.Share, contentDescription = null, modifier = Modifier.size(16.dp))
                            Spacer(modifier = Modifier.width(6.dp))
                            Text("Создать файл ОТМЕНА", fontWeight = FontWeight.SemiBold)
                        }
                    }
                }

                Spacer(modifier = Modifier.height(14.dp))

                // 4. ТАБЛИЦА РЕЗУЛЬТАТОВ РЕВИЗИИ
                Card(
                    modifier = Modifier.fillMaxSize(),
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = ThemeColors.CardBg),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Column(modifier = Modifier.fillMaxSize().padding(12.dp)) {
                        // Заголовок колонок
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .clip(RoundedCornerShape(6.dp))
                                .background(Color(0xFFF1F5F9))
                                .padding(vertical = 10.dp, horizontal = 12.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Text("ПОЕЗД", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = ThemeColors.HeaderSlate, modifier = Modifier.width(90.dp))
                            Text("СТАНЦИЯ", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = ThemeColors.HeaderSlate, modifier = Modifier.width(140.dp))
                            Text("ПО АФИШЕ", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = ThemeColors.HeaderSlate, modifier = Modifier.width(90.dp))
                            Text("ФАКТ / ТЕЛ.", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = ThemeColors.HeaderSlate, modifier = Modifier.width(100.dp))
                            Text("СТАТУС", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = ThemeColors.HeaderSlate, modifier = Modifier.width(130.dp))
                            Text("ДЕТАЛИ И ПРИМЕЧАНИЯ", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = ThemeColors.HeaderSlate, modifier = Modifier.weight(1f))
                        }

                        Divider(color = ThemeColors.BorderSubtle, modifier = Modifier.padding(vertical = 4.dp))

                        if (diffList.isEmpty()) {
                            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                Column(horizontalAlignment = Alignment.CenterHorizontally) {
                                    Icon(
                                        imageVector = Icons.Default.Search,
                                        contentDescription = null,
                                        tint = Color(0xFFCBD5E1),
                                        modifier = Modifier.size(48.dp)
                                    )
                                    Spacer(modifier = Modifier.height(8.dp))
                                    Text(
                                        text = "Нет данных для отображения. Загрузите файлы и нажмите «Сверить данные».",
                                        color = ThemeColors.TextMuted,
                                        fontSize = 14.sp
                                    )
                                }
                            }
                        } else {
                            LazyColumn(modifier = Modifier.fillMaxSize()) {
                                items(diffList) { item ->
                                    val (badgeBg, badgeText, statusBorder) = when (item.statusType) {
                                        StatusType.CANCELLED -> Triple(Color(0xFFFEE2E2), Color(0xFF991B1B), Color(0xFFFCA5A5))
                                        StatusType.MODIFIED -> Triple(Color(0xFFFEF3C7), Color(0xFF92400E), Color(0xFFFCD34D))
                                        StatusType.NORMAL -> Triple(Color(0xFFECFDF5), Color(0xFF065F46), Color(0xFFA7F3D0))
                                    }

                                    Row(
                                        modifier = Modifier
                                            .fillMaxWidth()
                                            .padding(vertical = 4.dp, horizontal = 4.dp)
                                            .clip(RoundedCornerShape(6.dp))
                                            .background(if (item.statusType == StatusType.CANCELLED) Color(0xFFFFF5F5) else Color.White)
                                            .border(1.dp, if (item.statusType == StatusType.CANCELLED) Color(0xFFFEE2E2) else Color(0xFFF1F5F9), RoundedCornerShape(6.dp))
                                            .padding(vertical = 8.dp, horizontal = 8.dp),
                                        verticalAlignment = Alignment.CenterVertically
                                    ) {
                                        // Номер поезда в строгом моноширинном стиле
                                        Text(
                                            text = item.trainNum,
                                            fontWeight = FontWeight.Bold,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            color = ThemeColors.HeaderSlate,
                                            modifier = Modifier.width(90.dp)
                                        )

                                        Text(
                                            text = item.station,
                                            fontSize = 13.sp,
                                            color = ThemeColors.TextMain,
                                            modifier = Modifier.width(140.dp)
                                        )

                                        Text(
                                            text = item.afishaTime,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            color = ThemeColors.TextMuted,
                                            modifier = Modifier.width(90.dp)
                                        )

                                        Text(
                                            text = item.actualTime,
                                            fontWeight = FontWeight.Bold,
                                            fontFamily = FontFamily.Monospace,
                                            fontSize = 13.sp,
                                            color = if (item.actualTime == "--:--") ThemeColors.DangerRed else ThemeColors.HeaderSlate,
                                            modifier = Modifier.width(100.dp)
                                        )

                                        // Бейдж статуса
                                        Surface(
                                            shape = RoundedCornerShape(4.dp),
                                            color = badgeBg,
                                            border = androidx.compose.foundation.BorderStroke(1.dp, statusBorder),
                                            modifier = Modifier.width(120.dp)
                                        ) {
                                            Text(
                                                text = item.status,
                                                color = badgeText,
                                                fontWeight = FontWeight.Bold,
                                                fontSize = 11.sp,
                                                modifier = Modifier.padding(horizontal = 6.dp, vertical = 3.dp)
                                            )
                                        }

                                        Spacer(modifier = Modifier.width(10.dp))

                                        Text(
                                            text = item.details,
                                            fontSize = 12.sp,
                                            color = ThemeColors.TextMain,
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

// Компонент карточки загрузки файлов
@Composable
fun ActionUploadCard(
    title: String,
    subtitle: String,
    icon: ImageVector,
    isLoaded: Boolean,
    color: Color,
    modifier: Modifier = Modifier,
    onClick: () -> Unit
) {
    Card(
        modifier = modifier,
        shape = RoundedCornerShape(10.dp),
        colors = CardDefaults.cardColors(containerColor = ThemeColors.CardBg),
        elevation = CardDefaults.cardElevation(1.dp)
    ) {
        Row(
            modifier = Modifier.fillMaxWidth().padding(12.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            Surface(
                shape = RoundedCornerShape(8.dp),
                color = if (isLoaded) color.copy(alpha = 0.12f) else Color(0xFFF1F5F9),
                modifier = Modifier.size(40.dp)
            ) {
                Box(contentAlignment = Alignment.Center) {
                    Icon(
                        imageVector = if (isLoaded) Icons.Default.Check else icon,
                        contentDescription = null,
                        tint = if (isLoaded) color else ThemeColors.TextMuted,
                        modifier = Modifier.size(20.dp)
                    )
                }
            }

            Spacer(modifier = Modifier.width(10.dp))

            Column(modifier = Modifier.weight(1f)) {
                Text(text = title, fontSize = 11.sp, fontWeight = FontWeight.Bold, color = ThemeColors.TextMuted)
                Text(
                    text = subtitle,
                    fontSize = 13.sp,
                    fontWeight = FontWeight.SemiBold,
                    color = ThemeColors.HeaderSlate,
                    maxLines = 1
                )
            }

            TextButton(
                onClick = onClick,
                contentPadding = PaddingValues(horizontal = 8.dp, vertical = 4.dp)
            ) {
                Text(if (isLoaded) "Заменить" else "Выбрать", fontSize = 12.sp, fontWeight = FontWeight.Bold, color = color)
            }
        }
    }
}

// Компонент индикатора статистики в шапке
@Composable
fun StatBadge(label: String, count: String, color: Color, bg: Color) {
    Surface(
        shape = RoundedCornerShape(6.dp),
        color = bg,
        border = androidx.compose.foundation.BorderStroke(1.dp, color.copy(alpha = 0.3f))
    ) {
        Row(
            modifier = Modifier.padding(horizontal = 10.dp, vertical = 4.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            Text(text = "$label: ", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = color)
            Text(text = count, fontSize = 13.sp, fontWeight = FontWeight.ExtraBold, color = color)
        }
    }
}

// Алгоритм синхронизации всех источников
fun compareAllSources(
    afisha: File,
    telegram: File?,
    vygruzki: List<File>,
    targetStation: String
): List<DiffItem> {
    val results = mutableListOf<DiffItem>()

    // 1. Читаем текст телеграммы
    val telText = buildString {
        telegram?.let { f ->
            try {
                if (f.name.endsWith(".docx", true)) {
                    FileInputStream(f).use { fis ->
                        XWPFDocument(fis).use { doc ->
                            doc.paragraphs.forEach { append(it.text).append("\n") }
                        }
                    }
                } else if (f.name.endsWith(".pdf", true)) {
                    Loader.loadPDF(f).use { doc ->
                        append(PDFTextStripper().getText(doc)).append("\n")
                    }
                }
            } catch (_: Exception) {}
        }
    }

    // 2. Читаем текст выгрузок
    val vygruzkiText = buildString {
        vygruzki.forEach { f ->
            try { append(f.readText()).append("\n") } catch (_: Exception) {}
        }
    }

    // 3. Читаем Афишу Excel по горизонтали со всех листов
    try {
        FileInputStream(afisha).use { fis ->
            val wb = WorkbookFactory.create(fis)
            val timeFmt = SimpleDateFormat("HH:mm")

            for (sheetIndex in 0 until wb.numberOfSheets) {
                val sheet = wb.getSheetAt(sheetIndex)

                var trainRowIndex = -1
                for (r in 0..6) {
                    val row = sheet.getRow(r) ?: continue
                    for (c in 1..25) {
                        val v = row.getCell(c)?.toString()?.trim() ?: ""
                        if (v.matches(Regex("\\d{4}(/\\d{4})?"))) {
                            trainRowIndex = r
                            break
                        }
                    }
                    if (trainRowIndex != -1) break
                }
                if (trainRowIndex == -1) continue

                var stationRowIndex = -1
                for (r in trainRowIndex + 1..sheet.lastRowNum) {
                    val row = sheet.getRow(r) ?: continue
                    val stationName = row.getCell(0)?.toString()?.trim() ?: ""
                    if (stationName.contains(targetStation, ignoreCase = true) ||
                        (targetStation == "Зеленоград-Крюково" && (stationName.contains("Крюково", true) || stationName.contains("Зеленоград", true)))
                    ) {
                        stationRowIndex = r
                        break
                    }
                }
                if (stationRowIndex == -1) continue

                val trainRow = sheet.getRow(trainRowIndex)
                val stRow = sheet.getRow(stationRowIndex)

                for (c in 1..trainRow.lastCellNum) {
                    val trainNum = trainRow.getCell(c.toInt())?.toString()?.trim()?.replace(".0", "") ?: ""
                    if (!trainNum.matches(Regex("\\d{4}(/\\d{4})?"))) continue

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

                    val numOnly = trainNum.split("/").first()
                    var status = "По графику"
                    var actualTime = afishaTime
                    var details = "Совпадает с афишей"
                    var type = StatusType.NORMAL

                    // Проверка отмены в телеграмме
                    if (telText.contains("П. НР $numOnly") || telText.contains("№ $numOnly") || telText.contains("№$numOnly")) {
                        val matchIdx = telText.indexOf(numOnly)
                        val otmSnippet = telText.substring((matchIdx - 30).coerceAtLeast(0), (matchIdx + 60).coerceAtMost(telText.length))
                        if (otmSnippet.contains("отмен", true) || otmSnippet.contains("ОТМЕН", true)) {
                            status = "ОТМЕНЕН"
                            actualTime = "--:--"
                            details = "Отменен телеграммой № 953"
                            type = StatusType.CANCELLED
                        }
                    }

                    // Проверка в выгрузках TXT
                    val vygRegex = Regex("№$numOnly.*?отправлением в (\\d{2}:\\d{2})\\s*-\\s*(.*?)(?=\\n|\$)")
                    val vygMatch = vygRegex.find(vygruzkiText)
                    if (vygMatch != null) {
                        val changeDesc = vygMatch.groupValues[2].trim()
                        if (changeDesc.contains("ОТМЕНЕН", true)) {
                            status = "ОТМЕНЕН"
                            actualTime = "--:--"
                            details = "Отменен по оперативной выгрузке"
                            type = StatusType.CANCELLED
                        } else if (changeDesc.contains("отправится в", true)) {
                            val newTimeMatch = Regex("отправится в (\\d{2}:\\d{2})").find(changeDesc)
                            val newTime = newTimeMatch?.groupValues?.get(1) ?: ""
                            status = "СДВИГ ГРАФИКА"
                            actualTime = newTime
                            details = changeDesc
                            type = StatusType.MODIFIED
                        }
                    }

                    if (type != StatusType.NORMAL) {
                        results.add(DiffItem(trainNum, targetStation, afishaTime, actualTime, status, details, type))
                    }
                }
            }
        }
    } catch (_: Exception) {}

    return results
}

// Выгрузка файла отчета "ОТМЕНА" на Рабочий стол
fun exportCancellationReport(items: List<DiffItem>, station: String) {
    val desktopDir = File(System.getProperty("user.home"), "Desktop")
    val outFile = File(desktopDir, "ОТМЕНА_${station.replace(" ", "_")}.txt")

    outFile.printWriter(Charsets.UTF_8).use { out ->
        out.println("$station Изменения в графике движения пригородных поездов")
        out.println("=".repeat(60))
        items.forEach { item ->
            if (item.statusType == StatusType.CANCELLED) {
                out.println("    №${item.trainNum} отправлением в ${item.afishaTime} - ОТМЕНЕН")
            } else if (item.statusType == StatusType.MODIFIED) {
                out.println("    №${item.trainNum} отправлением в ${item.afishaTime} - отправится в ${item.actualTime}. ${item.details}")
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
