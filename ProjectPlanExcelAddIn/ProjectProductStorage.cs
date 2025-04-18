using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

public class ProjectProductStorage
{
    public Dictionary<string, List<string>> ProjectsProducts { get; private set; } = new Dictionary<string, List<string>>();

    private readonly string _filePath;

    public ProjectProductStorage()
    {
        string folderPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ProjectPlan", "Data");
        _filePath = Path.Combine(folderPath, "projects_products_list.json");

        LoadFromFile();
    }

    private void LoadFromFile()
    {
        try
        {
            if (File.Exists(_filePath))
            {
                string json = File.ReadAllText(_filePath);
                ProjectsProducts = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json)
                                  ?? new Dictionary<string, List<string>>();
            }
            else
            {
                ProjectsProducts = new Dictionary<string, List<string>>();
            }
        }
        catch (Exception ex)
        {
            // Обработка ошибок
            System.Diagnostics.Debug.WriteLine("Ошибка при загрузке JSON: " + ex.Message);
            ProjectsProducts = new Dictionary<string, List<string>>();
        }
    }

    public void SaveToFile()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath));
            string json = JsonConvert.SerializeObject(ProjectsProducts, Formatting.Indented);
            File.WriteAllText(_filePath, json);
        }
        catch (Exception ex)
        {
            // Обработка ошибок
            System.Diagnostics.Debug.WriteLine("Ошибка при сохранении JSON: " + ex.Message);
        }
    }
}
