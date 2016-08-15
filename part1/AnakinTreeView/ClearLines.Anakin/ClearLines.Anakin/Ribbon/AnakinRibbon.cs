namespace ClearLines.Anakin.Ribbon
{
   using Microsoft.Office.Tools.Ribbon;

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
