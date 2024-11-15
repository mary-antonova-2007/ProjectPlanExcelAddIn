using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
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
        {
            throw new InvalidOperationException("API key or model is not configured.");
        }

        using (var client = new HttpClient())
        {
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {ApiKey}");

            int promptTokens = CountTokens(prompt);
            int maxResponseTokens = 4096 - promptTokens; // Общее количество токенов для модели минус токены запроса

            var requestBody = new
            {
                model = Model,
                messages = new[]
            {
                new {
                    role = "system",
                    content = content
                },
                new { role = "user", content = prompt },
                new { role = "user", content = $"Выбранный диапазон ячеек: {range}" }
            },
                max_tokens = maxResponseTokens,
                temperature = 0.7
            };

            var jsonBody = JsonConvert.SerializeObject(requestBody);
            var data = Encoding.UTF8.GetBytes(jsonBody);

            var request = (HttpWebRequest)WebRequest.Create("https://api.openai.com/v1/chat/completions");
            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            request.Headers.Add("Authorization", $"Bearer {ApiKey}");


            // Запись данных в поток запроса
            using (var stream = await request.GetRequestStreamAsync())
            {
                stream.Write(data, 0, data.Length);
            }
            // Получение ответа
            using (var response = await request.GetResponseAsync())
            using (var responseStream = response.GetResponseStream())
            using (var reader = new StreamReader(responseStream))
            {
                var responseText = await reader.ReadToEndAsync();
                dynamic result = JsonConvert.DeserializeObject(responseText);
                string responseContent = result.choices[0].message.content;
                return responseContent;
            }
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
}
