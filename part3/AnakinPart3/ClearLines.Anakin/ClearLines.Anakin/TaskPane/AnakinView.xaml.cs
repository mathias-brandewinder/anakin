namespace ClearLines.Anakin.TaskPane
{
   using System.Windows;
   using System.Windows.Controls;
   using ClearLines.Anakin.TaskPane.TreeView;

   /// <summary>
   /// Interaction logic for AnakinView.xaml
   /// </summary>
   public partial class AnakinView : UserControl
   {
      public AnakinView()
      {
         InitializeComponent();
      }

      internal AnakinViewModel ViewModel
      {
         get
         {
            return this.DataContext as AnakinViewModel;
         }
      }

      private void SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
      {
         var model = this.DataContext as AnakinViewModel;
         if (model != null)
         {
            var worksheetViewModel = e.NewValue as WorksheetViewModel;
            if (worksheetViewModel != null)
            {
               var worksheet = worksheetViewModel.Worksheet;
               model.SelectedWorksheet = worksheet;
            }

            else
            {
               model.SelectedWorksheet = null;
            }
         }
      }
   }
}
