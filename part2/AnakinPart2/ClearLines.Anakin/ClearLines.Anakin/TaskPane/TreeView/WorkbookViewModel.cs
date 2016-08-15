//-----------------------------------------------------------------------
// <copyright file="WorkbookViewModel.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane.TreeView
{
   using System.Collections.ObjectModel;
   using System.ComponentModel;
   using Excel = Microsoft.Office.Interop.Excel;

   /// <summary>
   /// WorkbookViewModel provides a representation of
   /// a workbook for the TreeView in the AnakinView control.
   /// It maintains a list of WorksheetViewModel, representing
   /// its Worksheets, and is responsible for updating the
   /// list when a sheet is added. 
   /// Because of the limitations of the events exposed by Excel,
   /// the removal of closed sheets, and the update of their names,
   /// is performed by ExcelViewModel.
   /// </summary>
   public class WorkbookViewModel : INotifyPropertyChanged
   {
      private readonly Excel.Workbook workbook;
      private readonly ObservableCollection<WorksheetViewModel> worksheetViewModels;
      private string name;
      private string author;

      internal WorkbookViewModel(Excel.Workbook workbook)
      {
         this.workbook = workbook;
         this.name = workbook.Name;
         this.author = workbook.Author;
         workbook.NewSheet += this.AddSheet;
         this.worksheetViewModels = new ObservableCollection<WorksheetViewModel>();
         var worksheets = workbook.Worksheets;
         foreach (var sheet in worksheets)
         {
            this.AddSheet(sheet);
         }
      }

      public event PropertyChangedEventHandler PropertyChanged;

      public string Name
      {
         get
         {
            return this.name;
         }

         set
         {
            if (value != this.name)
            {
               this.name = value;
               this.OnPropertyChanged("Name");
            }
         }
      }

      public string Author
      {
         get
         {
            return this.author;
         }

         set
         {
            if (value != this.author)
            {
               this.author = value;
               this.OnPropertyChanged("Author");
            }
         }
      }

      public string ImagePath
      {
         get
         {
            return "Treeview/Workbook.bmp";
         }
      }

      public ObservableCollection<WorksheetViewModel> Worksheets
      {
         get
         {
            return this.worksheetViewModels;
         }
      }

      internal Excel.Workbook Workbook
      {
         get
         {
            return this.workbook;
         }
      }

      internal void UpdateDisplayProperties()
      {
         this.Name = this.workbook.Name;
         this.Author = this.workbook.Author;
      }

      protected void OnPropertyChanged(string propertyName)
      {
         var handler = this.PropertyChanged;
         if (handler != null)
         {
            handler(this, new PropertyChangedEventArgs(propertyName));
         }
      }

      private void AddSheet(object newSheet)
      {
         var worksheet = newSheet as Excel.Worksheet;
         if (worksheet != null)
         {
            var worksheetViewModel = new WorksheetViewModel(worksheet);
            this.worksheetViewModels.Add(worksheetViewModel);
         }
      }
   }
}