// <copyright file="AnakinRibbon.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.Ribbon
{
    partial class AnakinRibbon
    {
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab2 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.AnakinGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.ShowAnakin = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab2.SuspendLayout();
            this.AnakinGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab2.ControlId.OfficeId = "TabReview";
            this.tab2.Groups.Add(this.AnakinGroup);
            this.tab2.Label = "TabReview";
            this.tab2.Name = "tab2";
            // 
            // AnakinGroup
            // 
            this.AnakinGroup.Items.Add(this.ShowAnakin);
            this.AnakinGroup.Label = "Anakin";
            this.AnakinGroup.Name = "AnakinGroup";
            // 
            // ShowAnakin
            // 
            this.ShowAnakin.Label = "Compare";
            this.ShowAnakin.Name = "ShowAnakin";
            this.ShowAnakin.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.ShowAnakin_Click);
            // 
            // AnakinRibbon
            // 
            this.Name = "AnakinRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab2);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.AnakinRibbon_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.AnakinGroup.ResumeLayout(false);
            this.AnakinGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AnakinGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowAnakin;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal AnakinRibbon AnakinRibbon
        {
            get { return this.GetRibbon<AnakinRibbon>(); }
        }
    }
}
