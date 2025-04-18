namespace ProjectPlanExcelAddIn
{
    partial class RibbonPlan : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonPlan()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabPlan = this.Factory.CreateRibbonTab();
            this.groupDates = this.Factory.CreateRibbonGroup();
            this.buttonAddDays = this.Factory.CreateRibbonButton();
            this.groupPlan = this.Factory.CreateRibbonGroup();
            this.buttonAutoPlan = this.Factory.CreateRibbonButton();
            this.ChatGPTGroup = this.Factory.CreateRibbonGroup();
            this.buttonGPTQuestion = this.Factory.CreateRibbonButton();
            this.buttonGPTSettings = this.Factory.CreateRibbonButton();
            this.groupMove = this.Factory.CreateRibbonGroup();
            this.buttonMoveRowUp = this.Factory.CreateRibbonButton();
            this.buttonMoveRowDown = this.Factory.CreateRibbonButton();
            this.buttonAddRowAbove = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonProductTimeReport = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.buttonUpdateProjectsProductsList = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.buttonClearAndLoadFromBook = this.Factory.CreateRibbonButton();
            this.tabPlan.SuspendLayout();
            this.groupDates.SuspendLayout();
            this.groupPlan.SuspendLayout();
            this.ChatGPTGroup.SuspendLayout();
            this.groupMove.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPlan
            // 
            this.tabPlan.Groups.Add(this.groupDates);
            this.tabPlan.Groups.Add(this.groupPlan);
            this.tabPlan.Groups.Add(this.ChatGPTGroup);
            this.tabPlan.Groups.Add(this.groupMove);
            this.tabPlan.Groups.Add(this.group1);
            this.tabPlan.Label = "Планирование";
            this.tabPlan.Name = "tabPlan";
            // 
            // groupDates
            // 
            this.groupDates.Items.Add(this.buttonAddDays);
            this.groupDates.Label = "Даты";
            this.groupDates.Name = "groupDates";
            // 
            // buttonAddDays
            // 
            this.buttonAddDays.Label = "Сдвинуть даты (дни)";
            this.buttonAddDays.Name = "buttonAddDays";
            this.buttonAddDays.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddDays_Click);
            // 
            // groupPlan
            // 
            this.groupPlan.Items.Add(this.buttonAutoPlan);
            this.groupPlan.Label = "Планирование";
            this.groupPlan.Name = "groupPlan";
            // 
            // buttonAutoPlan
            // 
            this.buttonAutoPlan.Label = "Авто планирование";
            this.buttonAutoPlan.Name = "buttonAutoPlan";
            this.buttonAutoPlan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAutoPlan_Click);
            // 
            // ChatGPTGroup
            // 
            this.ChatGPTGroup.Items.Add(this.buttonGPTQuestion);
            this.ChatGPTGroup.Items.Add(this.buttonGPTSettings);
            this.ChatGPTGroup.Label = "ChatGPT";
            this.ChatGPTGroup.Name = "ChatGPTGroup";
            // 
            // buttonGPTQuestion
            // 
            this.buttonGPTQuestion.Label = "Вопрос";
            this.buttonGPTQuestion.Name = "buttonGPTQuestion";
            this.buttonGPTQuestion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGPTQuestion_Click);
            // 
            // buttonGPTSettings
            // 
            this.buttonGPTSettings.Label = "Настройки";
            this.buttonGPTSettings.Name = "buttonGPTSettings";
            this.buttonGPTSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGPTSettings_Click);
            // 
            // groupMove
            // 
            this.groupMove.Items.Add(this.buttonMoveRowUp);
            this.groupMove.Items.Add(this.buttonMoveRowDown);
            this.groupMove.Items.Add(this.buttonAddRowAbove);
            this.groupMove.KeyTip = "W";
            this.groupMove.Label = "Перемещение строк";
            this.groupMove.Name = "groupMove";
            // 
            // buttonMoveRowUp
            // 
            this.buttonMoveRowUp.KeyTip = "W1";
            this.buttonMoveRowUp.Label = "Строки вверх";
            this.buttonMoveRowUp.Name = "buttonMoveRowUp";
            this.buttonMoveRowUp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveRowsUp);
            // 
            // buttonMoveRowDown
            // 
            this.buttonMoveRowDown.KeyTip = "S1";
            this.buttonMoveRowDown.Label = "Строки вниз";
            this.buttonMoveRowDown.Name = "buttonMoveRowDown";
            this.buttonMoveRowDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveRowsDown);
            // 
            // buttonAddRowAbove
            // 
            this.buttonAddRowAbove.KeyTip = "Q1";
            this.buttonAddRowAbove.Label = "Добавить строку выше";
            this.buttonAddRowAbove.Name = "buttonAddRowAbove";
            this.buttonAddRowAbove.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertRowAbove);
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonProductTimeReport);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.buttonUpdateProjectsProductsList);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.buttonClearAndLoadFromBook);
            this.group1.KeyTip = "W";
            this.group1.Label = "Отчёты";
            this.group1.Name = "group1";
            // 
            // buttonProductTimeReport
            // 
            this.buttonProductTimeReport.KeyTip = "W1";
            this.buttonProductTimeReport.Label = "Отчет по изделиям (файлы)";
            this.buttonProductTimeReport.Name = "buttonProductTimeReport";
            this.buttonProductTimeReport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonProductTimeReport_Click);
            // 
            // button1
            // 
            this.button1.KeyTip = "W1";
            this.button1.Label = "Отчет по изделиям (папка подпапки)";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateReport_Click);
            // 
            // buttonUpdateProjectsProductsList
            // 
            this.buttonUpdateProjectsProductsList.Label = "Загрузить новые изделия и проекты";
            this.buttonUpdateProjectsProductsList.Name = "buttonUpdateProjectsProductsList";
            this.buttonUpdateProjectsProductsList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUpdateProjectsProductsList_Click);
            // 
            // button2
            // 
            this.button2.Label = "Сохранить проекты изделия в книгу";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // buttonClearAndLoadFromBook
            // 
            this.buttonClearAndLoadFromBook.Label = "Загрузить данные проектов из книги и очистить старые";
            this.buttonClearAndLoadFromBook.Name = "buttonClearAndLoadFromBook";
            this.buttonClearAndLoadFromBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonClearAndLoadFromBook_Click);
            // 
            // RibbonPlan
            // 
            this.Name = "RibbonPlan";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabPlan);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonPlan_Load);
            this.tabPlan.ResumeLayout(false);
            this.tabPlan.PerformLayout();
            this.groupDates.ResumeLayout(false);
            this.groupDates.PerformLayout();
            this.groupPlan.ResumeLayout(false);
            this.groupPlan.PerformLayout();
            this.ChatGPTGroup.ResumeLayout(false);
            this.ChatGPTGroup.PerformLayout();
            this.groupMove.ResumeLayout(false);
            this.groupMove.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabPlan;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddDays;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPlan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAutoPlan;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ChatGPTGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGPTQuestion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGPTSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMove;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMoveRowUp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMoveRowDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddRowAbove;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonProductTimeReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUpdateProjectsProductsList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonClearAndLoadFromBook;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonPlan RibbonPlan
        {
            get { return this.GetRibbon<RibbonPlan>(); }
        }
    }
}
