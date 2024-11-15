using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectPlanExcelAddIn
{
    public partial class GPTSettingsForm : Form
    {
        public string SelectedModel { get; set; }
        public string ApiKey { get; set; }
        private List<string> modelOptions = new List<string>
        {
            "chatgpt-4o-latest",
            "gpt-4-turbo",
            "gpt-4-turbo-2024-04-09",
            "gpt-4",
            "gpt-4-32k",
            "gpt-4-0125-preview",
            "gpt-4-1106-preview",
            "gpt-4-vision-preview",
            "gpt-3.5-turbo-0125",
            "gpt-3.5-turbo-instruct",
            "gpt-3.5-turbo-1106",
            "gpt-3.5-turbo-0613",
            "gpt-3.5-turbo-16k-0613",
            "gpt-3.5-turbo-0301",
            "davinci-002",
            "babbage-002"
        };
        public GPTSettingsForm()
        {
            InitializeComponent();
        }
        private void buttonSave_Click(object sender, EventArgs e)
        {
            SelectedModel = comboBoxModels.SelectedItem.ToString();
            ApiKey = textBoxApiKey.Text;
            SaveSettings();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        private void SaveSettings()
        {
            var config = new
            {
                ApiKey = this.ApiKey,
                SelectedModel = this.SelectedModel
            };

            string configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Planer",
                "GPTManager.json"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(configPath));
            File.WriteAllText(configPath, JsonConvert.SerializeObject(config));
        }
        private void SettingsForm_Load(object sender, EventArgs e)
        {
            LoadSettings();
            comboBoxModels.SelectedItem = SelectedModel;
            textBoxApiKey.Text = ApiKey;
        }

        private void LoadSettings()
        {
            comboBoxModels.Items.Clear();
            foreach (var model in modelOptions) { comboBoxModels.Items.Add(model); }
            
            string configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Planer",
                "GPTManager.json"
            );

            if (File.Exists(configPath))
            {
                var config = JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(configPath));
                SelectedModel = config.SelectedModel;
                ApiKey = config.ApiKey;
            }
        }

    }
}
