﻿namespace ProjectPlanExcelAddIn
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
            this.tabPlan.SuspendLayout();
            this.groupDates.SuspendLayout();
            this.groupPlan.SuspendLayout();
            this.ChatGPTGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPlan
            // 
            this.tabPlan.Groups.Add(this.groupDates);
            this.tabPlan.Groups.Add(this.groupPlan);
            this.tabPlan.Groups.Add(this.ChatGPTGroup);
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
    }

    partial class ThisRibbonCollection
    {
        internal RibbonPlan RibbonPlan
        {
            get { return this.GetRibbon<RibbonPlan>(); }
        }
    }
}
