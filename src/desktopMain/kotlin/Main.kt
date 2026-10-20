import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
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
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.awt.FileDialog
import java.awt.Frame
import java.io.File
import java.io.FileInputStream

// Модель рейса / нитки поезда
data class ScheduleItem(
    val trainNumber: String,
    val route: String,
    val departureTime: String,
    val status: String,
    val note: String,
    val statusType: StatusType = StatusType.NORMAL
)

enum class StatusType {
    NORMAL, MODIFIED, CANCELLED
}

@Composable
fun App() {
    var excelFile by remember { mutableStateOf<File?>(null) }
    var telegramFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var scheduleList by remember { mutableStateOf<List<ScheduleItem>>(emptyList()) }
    var isProcessing by remember { mutableStateOf(false) }
    var logMessage by remember { mutableStateOf("Выберите файл афиши (Excel) и телеграммы для начала сверки") }

    val coroutineScope = rememberCoroutineScope()

    // Функция вызова системного окна выбора файла
    fun selectFile(title: String, extensions: Array<String>, multiple: Boolean = false): List<File> {
        val dialog = FileDialog(null as Frame?, title, FileDialog.LOAD).apply {
            isMultipleMode = multiple
            isVisible = true
        }
        val files = dialog.files.toList()
        dialog.dispose()
        return files
    }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = Color(0xFFF7F9FC)) {
            Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
                // Заголовок
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    colors = CardDefaults.cardColors(containerColor = Color.White),
                    elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                ) {
                    Column(modifier = Modifier.padding(16.dp)) {
                        Text(
                            text = "Ревизор расписаний с приоритетом телеграмм",
                            fontSize = 20.sp,
                            fontWeight = FontWeight.Bold,
                            color = Color(0xFF1E293B)
                        )
                        Spacer(modifier = Modifier.height(8.dp))
                        Text(
                            text = logMessage,
                            fontSize = 13.sp,
                            color = if (isProcessing) Color(0xFF2563EB) else Color(0xFF64748B)
                        )
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))

                // Панель кнопок выбора файлов и запуска
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(8.dp),
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Button(
                        onClick = {
                            val picked = selectFile("Выберите афишу расписания", arrayOf(".xlsx", ".xls"))
                            if (picked.isNotEmpty()) {
                                excelFile = picked.first()
                                logMessage = "Афиша выбрана: ${excelFile?.name}"
                            }
                        },
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF0F766E))
                    ) {
                        Text(if (excelFile == null) "1. Выбрать афишу (.xlsx)" else "✓ ${excelFile?.name}")
                    }

                    Button(
                        onClick = {
                            val picked = selectFile("Выберите телеграммы и выгрузки", arrayOf(".pdf", ".docx", ".txt"), multiple = true)
                            if (picked.isNotEmpty()) {
                                telegramFiles = picked
                                logMessage = "Выбрано телеграмм/выгрузок: ${picked.size}"
                            }
                        },
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF0369A1))
                    ) {
                        Text(if (telegramFiles.isEmpty()) "2. Добавить телеграммы (PDF/Word/TXT)" else "✓ Телеграмм: ${telegramFiles.size}")
                    }

                    Spacer(modifier = Modifier.weight(1f))

                    Button(
                        onClick = {
                            if (excelFile == null) {
                                logMessage = "Ошибка: сначала выберите файл афиши Excel!"
                                return@Button
                            }
                            coroutineScope.launch {
                                isProcessing = true
                                logMessage = "Выполняется парсинг и сверка с телеграммами..."
                                
                                val parsed = withContext(Dispatchers.IO) {
                                    runVerification(excelFile!!, telegramFiles)
                                }
                                
                                scheduleList = parsed
                                isProcessing = false
                                logMessage = "Сверка завершена! Обработано поездов: ${parsed.size}"
                            }
                        },
                        enabled = !isProcessing && excelFile != null,
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF4338CA))
                    ) {
                        Text(if (isProcessing) "Обработка..." else "Выполнить сверку")
                    }
                }

                Spacer(modifier = Modifier.height(12.dp))

                // Таблица результатов
                Card(
                    modifier = Modifier.fillMaxSize(),
                    colors = CardDefaults.cardColors(containerColor = Color.White),
                    elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                ) {
                    Column(modifier = Modifier.fillMaxSize().padding(12.dp)) {
                        // Заголовки таблицы
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .background(Color(0xFFE2E8F0), RoundedCornerShape(4.dp))
                                .padding(vertical = 8.dp, horizontal = 12.dp)
                        ) {
                            Text("Поезд", fontWeight = FontWeight.SemiBold, modifier = Modifier.width(100.dp))
                            Text("Маршрут", fontWeight = FontWeight.SemiBold, modifier = Modifier.weight(1.5f))
                            Text("Время отпр.", fontWeight = FontWeight.SemiBold, modifier = Modifier.width(110.dp))
                            Text("Статус", fontWeight = FontWeight.SemiBold, modifier = Modifier.width(130.dp))
                            Text("Примечание / Телеграмма", fontWeight = FontWeight.SemiBold, modifier = Modifier.weight(2f))
                        }

                        Divider(modifier = Modifier.padding(vertical = 4.dp), color = Color(0xFFCBD5E1))

                        if (scheduleList.isEmpty()) {
                            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                Text("Нет данных для отображения. Загрузите файлы и нажмите «Выполнить сверку».", color = Color.Gray)
                            }
                        } else {
                            LazyColumn(modifier = Modifier.fillMaxSize()) {
                                items(scheduleList) { item ->
                                    val (bgColor, textColor) = when (item.statusType) {
                                        StatusType.CANCELLED -> Color(0xFFFFE4E6) to Color(0xFF9F1239)
                                        StatusType.MODIFIED -> Color(0xFFFEF3C7) to Color(0xFF92400E)
                                        StatusType.NORMAL -> Color.Transparent to Color(0xFF1E293B)
                                    }

                                    Row(
                                        modifier = Modifier
                                            .fillMaxWidth()
                                            .background(bgColor, RoundedCornerShape(4.dp))
                                            .padding(vertical = 8.dp, horizontal = 12.dp),
                                        verticalAlignment = Alignment.CenterVertically
                                    ) {
                                        Text(item.trainNumber, fontWeight = FontWeight.Medium, modifier = Modifier.width(100.dp))
                                        Text(item.route, modifier = Modifier.weight(1.5f))
                                        Text(item.departureTime, modifier = Modifier.width(110.dp))
                                        Text(item.status, color = textColor, fontWeight = FontWeight.SemiBold, modifier = Modifier.width(130.dp))
                                        Text(item.note, color = textColor, modifier = Modifier.weight(2f))
                                    }
                                    Divider(color = Color(0xFFF1F5F9))
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

// Фоновая логика чтения документов и применения изменений телеграмм
suspend fun runVerification(excelFile: File, telegrams: List<File>): List<ScheduleItem> {
    val items = mutableListOf<ScheduleItem>()

    // 1. Извлекаем сырой текст из всех приложенных телеграмм (PDF, Word, TXT)
    val telegramTexts = buildString {
        for (file in telegrams) {
            try {
                when {
                    file.name.endsWith(".pdf", ignoreCase = true) -> {
                        Loader.loadPDF(file).use { doc ->
                            append(PDFTextStripper().getText(doc)).append("\n")
                        }
                    }
                    file.name.endsWith(".docx", ignoreCase = true) -> {
                        FileInputStream(file).use { fis ->
                            XWPFDocument(fis).use { doc ->
                                for (p in doc.paragraphs) {
                                    append(p.text).append("\n")
                                }
                            }
                        }
                    }
                    file.name.endsWith(".txt", ignoreCase = true) -> {
                        append(file.readText()).append("\n")
                    }
                }
            } catch (_: Exception) {}
        }
    }

    // 2. Читаем Excel-афишу расписания
    try {
        FileInputStream(excelFile).use { fis ->
            val workbook = WorkbookFactory.create(fis)
            val sheet = workbook.getSheetAt(0)

            for (rowIndex in 1..sheet.lastRowNum) {
                val row = sheet.getRow(rowIndex) ?: continue
                val trainNum = row.getCell(0)?.toString()?.trim()?.replace(".0", "") ?: ""
                val route = row.getCell(1)?.toString()?.trim() ?: "Маршрут не указан"
                val depTime = row.getCell(2)?.toString()?.trim() ?: "--:--"

                if (trainNum.isNotEmpty()) {
                    var status = "По расписанию"
                    var note = "Без изменений"
                    var statusType = StatusType.NORMAL

                    // Проверяем наличие номера поезда в тексте телеграмм
                    if (telegramTexts.contains(trainNum)) {
                        // Поиск контекста вокруг номера
                        val matchIndex = telegramTexts.indexOf(trainNum)
                        val snippet = telegramTexts.substring(
                            (matchIndex - 40).coerceAtLeast(0),
                            (matchIndex + 80).coerceAtMost(telegramTexts.length)
                        ).lowercase()

                        if (snippet.contains("отмен") || snippet.contains("отм")) {
                            status = "ОТМЕНЁН"
                            note = "Телеграмма: рейс отменяется"
                            statusType = StatusType.CANCELLED
                        } else if (snippet.contains("измен") || snippet.contains("расписан")) {
                            status = "ИЗМЕНЁН"
                            note = "Телеграмма: скорректирован график"
                            statusType = StatusType.MODIFIED
                        } else {
                            status = "УПОМЯНУТ"
                            note = "Найден в тексте телеграммы"
                            statusType = StatusType.MODIFIED
                        }
                    }

                    items.add(ScheduleItem(trainNum, route, depTime, status, note, statusType))
                }
            }
        }
    } catch (_: Exception) {}

    return items
}

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "ScheduleVerifier — Ревизор расписаний с приоритетом телеграмм"
    ) {
        App()
    }
}
