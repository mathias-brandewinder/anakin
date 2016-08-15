// <copyright file="AnakinRibbon.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.Ribbon
{
   using Microsoft.Office.Tools.Ribbon;

   /// <summary>
   /// The AnakinRibbon extends the Excel Ribbon,
   /// adding to the Review Tab a button to show
   /// the Custom Task Pane for the add-in.
   /// </summary>
   public partial class AnakinRibbon : OfficeRibbon
    {
        public AnakinRibbon()
        {
            InitializeComponent();
        }

        private void AnakinRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void ShowAnakin_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPane.Visible = true;
        }
    }
}
