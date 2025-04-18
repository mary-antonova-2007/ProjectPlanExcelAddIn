using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using ProjectPlanExcelAddIn;

public class GPTManager
{
    public string ApiKey { get; private set; }
    public string Model { get; private set; }

    private string content = "Ты работаешь в надстройке для Excel." +
        " Ответ должен содержать JSON-команды для выполнения действий в Excel, если сообщение требует таких команд." +
        " Формат ответа: [{ \"action\": \"write\", \"cell\": \"A1\", \"value\": \"Москва\" }, {...}]." +
        " Если нужен диапазон, то указывается cell, например, \"A1:A10\"." +
        "Если команд нет то просто ответь на вопрос текстом. В конце текста не должно быть никих выводов и комментариев. Только решение задачи или результат для вставки в excel";
    public GPTManager()
    {
        LoadSettings();
    }
    private void LoadSettings()
    {
        string configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Planer",
            "GPTManager.json"
        );

        if (File.Exists(configPath))
        {
            var config = JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(configPath));
            ApiKey = config.ApiKey;
            Model = config.SelectedModel;
        }
    }
    public async Task<string> GetResponseAsync(string prompt, string range)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        if (string.IsNullOrEmpty(ApiKey) || string.IsNullOrEmpty(Model))
            throw new InvalidOperationException("API key or model is not configured.");

        try
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ApiKey);

                int promptTokens = CountTokens(prompt);
                int maxResponseTokens = 4096 - promptTokens;

                var requestBody = new
                {
                    model = Model,
                    messages = new[]
                    {
                    new { role = "system", content = content },
                    new { role = "user", content = prompt },
                    new { role = "user", content = $"Выбранный диапазон ячеек: {range}" }
                },
                    max_tokens = maxResponseTokens,
                    temperature = 0.7
                };

                string json = JsonConvert.SerializeObject(requestBody);
                var httpContent = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("https://api.openai.com/v1/chat/completions", httpContent);
                string responseText = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    return $"Ошибка: {response.StatusCode}\n{responseText}";
                }

                dynamic result = JsonConvert.DeserializeObject(responseText);
                return result.choices[0].message.content;
            }
        }
        catch (Exception ex)
        {
            return $"Ошибка: {ex.Message}";
        }
    }
    public void ExecuteCommands(string responseContent)
    {
        try
        {
            // Пробуем распарсить ответ как JSON-команды
            var commands = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(responseContent);

            // Если десериализация прошла успешно, выполняем команды
            if (commands != null)
            {
                foreach (var command in commands)
                {
                    if (command.TryGetValue("action", out string action) && action == "write" &&
                        command.TryGetValue("cell", out string cell) &&
                        command.TryGetValue("value", out string value))
                    {
                        // Запись значения в указанную ячейку
                        var excelCell = Globals.ThisAddIn.Application.Range[cell];
                        excelCell.Value2 = value;
                    }
                }
            }
            else
            {
                // Если не удалось распарсить как JSON, записываем ответ как текст в активную ячейку
                var activeCell = Globals.ThisAddIn.Application.ActiveCell;
                activeCell.Value2 = responseContent;
            }
        }
        catch (JsonException)
        {
            // В случае ошибки при десериализации, считаем, что это текстовый ответ, и пишем его в активную ячейку
            var activeCell = Globals.ThisAddIn.Application.ActiveCell;
            activeCell.Value2 = responseContent;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при выполнении команды: {ex.Message}");
        }
    }

    int CountTokens(string input)
    {
        // Примерный подсчет токенов: каждый 4-5 символов можно считать за 1 токен.
        return input.Length / 4;
    }
    public async Task<string> DescribeImageAsync(string filePath)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        try
        {
            using (var client = new HttpClient())
            {
                // Устанавливаем авторизацию
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ApiKey);

                // Загружаем изображение и кодируем в Base64
                byte[] imageData = File.ReadAllBytes(filePath);
                string base64Image = Convert.ToBase64String(imageData);

                // Формируем тело запроса
                var requestBody = new
                {
                    model = "gpt-4o", // Замените модель на доступную
                    messages = new[]
                    {
                    new
                    {
                        role = "user",
                        content = $"Опиши изображение: [data:image/jpeg;base64,{base64Image}]"
                    }
                },
                    max_tokens = 1000 // Устанавливаем ограничение на количество токенов в ответе
                };

                // Сериализуем тело запроса
                string jsonContent = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                // Отправляем POST-запрос
                HttpResponseMessage response = await client.PostAsync("https://api.openai.com/v1/chat/completions", content);

                if (response.IsSuccessStatusCode)
                {
                    // Читаем и возвращаем тело ответа
                    string responseBody = await response.Content.ReadAsStringAsync();
                    return responseBody;
                }
                else
                {
                    // В случае ошибки возвращаем детали
                    string errorDetails = await response.Content.ReadAsStringAsync();
                    return $"Ошибка при получении ответа: {response.ReasonPhrase}. Детали: {errorDetails}";
                }
            }
        }
        catch (Exception ex)
        {
            // Обработка исключений
            return $"Произошла ошибка: {ex.Message}";
        }
    }
}
