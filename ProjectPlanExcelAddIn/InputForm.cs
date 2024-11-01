using System;
using System.Windows.Forms;

namespace ProjectPlanExcelAddIn
{
    public partial class InputForm : Form
    {
        public string TextBoxData {
            get { return textBoxData.Text; }
            set { textBoxData.Text = value; }
        }
        public string LabelInfo
        {
            get { return labelInfo.Text; }
            set { labelInfo.Text = value; }
        }
        public InputForm()
        {
            InitializeComponent();

            // Устанавливаем обработчики событий для кнопок
            buttonOk.Click += ButtonOk_Click;
            buttonCancel.Click += ButtonCancel_Click;
        }

        // Метод для обработки кнопки OK
        private void ButtonOk_Click(object sender, EventArgs e)
        {
            // Закрываем форму и возвращаем DialogResult.OK
            DialogResult = DialogResult.OK;
            Close();
        }

        // Метод для обработки кнопки Cancel
        private void ButtonCancel_Click(object sender, EventArgs e)
        {
            // Закрываем форму и возвращаем DialogResult.Cancel
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
