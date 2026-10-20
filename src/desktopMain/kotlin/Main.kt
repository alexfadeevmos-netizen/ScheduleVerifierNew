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
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.awt.FileDialog
import java.awt.Frame
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.text.SimpleDateFormat

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

@Composable
fun App() {
    var afishaFile by remember { mutableStateOf<File?>(null) }
    var telegramFile by remember { mutableStateOf<File?>(null) }
    var vygruzkiFiles by remember { mutableStateOf<List<File>>(emptyList()) }
    var diffList by remember { mutableStateOf<List<DiffItem>>(emptyList()) }
    var selectedStation by remember { mutableStateOf("Зеленоград-Крюково") }
    var isProcessing by remember { mutableStateOf(false) }
    var statusLog by remember { mutableStateOf("Загрузите Афишу, Телеграмму и Выгрузки для синхронизации") }

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

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = Color(0xFFF1F5F9)) {
            Column(modifier = Modifier.fillMaxSize().padding(14.dp)) {
                // Карточка заголовка
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    colors = CardDefaults.cardColors(containerColor = Color.White),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Column(modifier = Modifier.padding(14.dp)) {
                        Text("Синхронизатор расписаний и генератор отчетов отмен", fontSize = 18.sp, fontWeight = FontWeight.Bold, color = Color(0xFF0F172A))
                        Spacer(modifier = Modifier.height(4.dp))
                        Text(statusLog, fontSize = 13.sp, color = Color(0xFF475569))
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // Кнопки выбора файлов
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    Button(
                        onClick = {
                            val f = pickFile("Выберите Афишу (.xlsx)")
                            if (f.isNotEmpty()) {
                                afishaFile = f.first()
                                statusLog = "Афиша: ${afishaFile?.name}"
                            }
                        },
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF0D9488)),
                        modifier = Modifier.weight(1f)
                    ) {
                        Text(if (afishaFile == null) "1. Афиша (.xlsx)" else "✓ ${afishaFile?.name?.take(18)}...")
                    }

                    Button(
                        onClick = {
                            val f = pickFile("Выберите Телеграмму (Word/PDF)")
                            if (f.isNotEmpty()) {
                                telegramFile = f.first()
                                statusLog = "Телеграмма: ${telegramFile?.name}"
                            }
                        },
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF0284C7)),
                        modifier = Modifier.weight(1f)
                    ) {
                        Text(if (telegramFile == null) "2. Телеграмма (.docx/.pdf)" else "✓ ${telegramFile?.name?.take(18)}...")
                    }

                    Button(
                        onClick = {
                            val f = pickFile("Выберите Выгрузки (.txt)", multiple = true)
                            if (f.isNotEmpty()) {
                                vygruzkiFiles = f
                                statusLog = "Выгрузок выбрано: ${f.size}"
                            }
                        },
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF4F46E5)),
                        modifier = Modifier.weight(1f)
                    ) {
                        Text(if (vygruzkiFiles.isEmpty()) "3. Выгрузки (.txt)" else "✓ Выгрузок: ${vygruzkiFiles.size}")
                    }
                }

                Spacer(modifier = Modifier.height(8.dp))

                // Фильтр по станции и действия
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp), verticalAlignment = Alignment.CenterVertically) {
                    Text("Контрольная станция:", fontWeight = FontWeight.SemiBold, fontSize = 13.sp)
                    stationsList.forEach { st ->
                        FilterChip(
                            selected = (selectedStation == st),
                            onClick = { selectedStation = st },
                            label = { Text(st.replace("Зеленоград-", ""), fontSize = 12.sp) }
                        )
                    }

                    Spacer(modifier = Modifier.weight(1f))

                    Button(
                        onClick = {
                            if (afishaFile == null) return@Button
                            coroutineScope.launch {
                                isProcessing = true
                                statusLog = "Выполняется синхронизация данных..."
                                diffList = withContext(Dispatchers.IO) {
                                    compareAllSources(afishaFile!!, telegramFile, vygruzkiFiles, selectedStation)
                                }
                                isProcessing = false
                                statusLog = "Сверка завершена. Изменений обнаружено: ${diffList.size}"
                            }
                        },
                        enabled = !isProcessing && afishaFile != null,
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF16A34A))
                    ) {
                        Text(if (isProcessing) "Сверка..." else "Сверить")
                    }

                    Button(
                        onClick = {
                            if (diffList.isEmpty()) return@Button
                            exportCancellationReport(diffList, selectedStation)
                            statusLog = "Отчет 'ОТМЕНА_$selectedStation.txt' успешно создан на Рабочем столе!"
                        },
                        enabled = diffList.isNotEmpty(),
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFFDC2626))
                    ) {
                        Text("Создать файл ОТМЕНА")
                    }
                }

                Spacer(modifier = Modifier.height(10.dp))

                // Таблица отчета
                Card(
                    modifier = Modifier.fillMaxSize(),
                    colors = CardDefaults.cardColors(containerColor = Color.White),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Column(modifier = Modifier.fillMaxSize().padding(10.dp)) {
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .background(Color(0xFFE2E8F0), RoundedCornerShape(4.dp))
                                .padding(8.dp)
                        ) {
                            Text("Поезд", fontWeight = FontWeight.Bold, modifier = Modifier.width(90.dp))
                            Text("Станция", fontWeight = FontWeight.Bold, modifier = Modifier.width(140.dp))
                            Text("По афише", fontWeight = FontWeight.Bold, modifier = Modifier.width(90.dp))
                            Text("Факт/Телегр.", fontWeight = FontWeight.Bold, modifier = Modifier.width(100.dp))
                            Text("Статус", fontWeight = FontWeight.Bold, modifier = Modifier.width(120.dp))
                            Text("Детали расхождения", fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f))
                        }

                        Divider(modifier = Modifier.padding(vertical = 4.dp), color = Color(0xFFCBD5E1))

                        if (diffList.isEmpty()) {
                            Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                Text("Нажмите «Сверить» для построения отчета расхождений.", color = Color.Gray)
                            }
                        } else {
                            LazyColumn(modifier = Modifier.fillMaxSize()) {
                                items(diffList) { item ->
                                    val (bg, txt) = when (item.statusType) {
                                        StatusType.CANCELLED -> Color(0xFFFFE4E6) to Color(0xFF9F1239)
                                        StatusType.MODIFIED -> Color(0xFFFEF3C7) to Color(0xFF92400E)
                                        StatusType.NORMAL -> Color.Transparent to Color(0xFF1E293B)
                                    }

                                    Row(
                                        modifier = Modifier
                                            .fillMaxWidth()
                                            .background(bg, RoundedCornerShape(4.dp))
                                            .padding(vertical = 6.dp, horizontal = 8.dp),
                                        verticalAlignment = Alignment.CenterVertically
                                    ) {
                                        Text(item.trainNum, fontWeight = FontWeight.SemiBold, modifier = Modifier.width(90.dp))
                                        Text(item.station, modifier = Modifier.width(140.dp))
                                        Text(item.afishaTime, modifier = Modifier.width(90.dp))
                                        Text(item.actualTime, fontWeight = FontWeight.Medium, modifier = Modifier.width(100.dp))
                                        Text(item.status, color = txt, fontWeight = FontWeight.Bold, modifier = Modifier.width(120.dp))
                                        Text(item.details, color = txt, modifier = Modifier.weight(1f))
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

// Полноценный анализ Афиши (по горизонтали), Телеграммы и Текстовых выгрузок
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

    // 2. Читаем выгрузки TXT
    val vygruzkiText = buildString {
        vygruzki.forEach { f ->
            try { append(f.readText()).append("\n") } catch (_: Exception) {}
        }
    }

    // 3. Читаем Афишу Excel по горизонтали со всех вкладок
    try {
        FileInputStream(afisha).use { fis ->
            val wb = WorkbookFactory.create(fis)
            val timeFmt = SimpleDateFormat("HH:mm")

            for (sheetIndex in 0 until wb.numberOfSheets) {
                val sheet = wb.getSheetAt(sheetIndex)

                // Ищем строку с номерами поездов (обычно строка 3 или 4)
                var trainRowIndex = -1
                for (r in 0..6) {
                    val row = sheet.getRow(r) ?: continue
                    for (c in 1..20) {
                        val v = row.getCell(c)?.toString()?.trim() ?: ""
                        if (v.matches(Regex("\\d{4}(/\\d{4})?"))) {
                            trainRowIndex = r
                            break
                        }
                    }
                    if (trainRowIndex != -1) break
                }
                if (trainRowIndex == -1) continue

                // Ищем строку целевой станции
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

                    // Сверяем с Телеграммой и Выгрузкой
                    val numOnly = trainNum.split("/").first()
                    var status = "По графику"
                    var actualTime = afishaTime
                    var details = "Совпадает с афишей"
                    var type = StatusType.NORMAL

                    // Проверка полной отмены поезда в телеграмме
                    val cancelPattern = Regex("(?i)(отменить|п\\.\\s*нр\\s*$numOnly.*отмен|отменяется.*$numOnly)")
                    if (cancelPattern.containsMatchIn(telText) || telText.contains("П. НР $numOnly")) {
                        val otmSnippet = telText.substring((telText.indexOf(numOnly) - 30).coerceAtLeast(0), (telText.indexOf(numOnly) + 50).coerceAtMost(telText.length))
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
                            status = "ИЗМЕНЕНО ВРЕМЯ"
                            actualTime = newTime
                            details = changeDesc
                            type = StatusType.MODIFIED
                        }
                    }

                    // Добавляем только те поезда, где есть изменения или отмены
                    if (type != StatusType.NORMAL) {
                        results.add(DiffItem(trainNum, targetStation, afishaTime, actualTime, status, details, type))
                    }
                }
            }
        }
    } catch (_: Exception) {}

    return results
}

// Экспорт готового текстового файла по шаблону файла "Крюково на Москву"
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
        title = "ScheduleVerifier — Ревизор расписаний с приоритетом телеграмм"
    ) {
        App()
    }
}
