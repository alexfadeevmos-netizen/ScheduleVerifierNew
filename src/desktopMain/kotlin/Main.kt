import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.unit.dp
import androidx.compose.ui.window.Window
import androidx.compose.ui.window.application

@Composable
fun App() {
    var status by remember { mutableStateOf("Готов к работе") }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize()) {
            Column(
                modifier = Modifier.fillMaxSize().padding(24.dp),
                verticalArrangement = Arrangement.Center,
                horizontalAlignment = Alignment.CenterHorizontally
            ) {
                Text(
                    text = "Ревизор расписаний с приоритетом телеграмм",
                    style = MaterialTheme.typography.headlineSmall
                )
                Spacer(modifier = Modifier.height(16.dp))
                Button(onClick = { status = "Загрузка файлов..." }) {
                    Text("Выбрать файл расписания")
                }
                Spacer(modifier = Modifier.height(12.dp))
                Text(text = status, style = MaterialTheme.typography.bodyMedium)
            }
        }
    }
}

fun main() = application {
    Window(
        onCloseRequest = ::exitApplication,
        title = "ScheduleVerifier"
    ) {
        App()
    }
}
