//-----------------------------------------------------------------------
// <copyright file="WorksheetViewModel.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane.TreeView
{
   using System.ComponentModel;
   using Excel = Microsoft.Office.Interop.Excel;

   public class WorksheetViewModel : INotifyPropertyChanged
   {
      private Excel.Worksheet worksheet;
      private string name;

      public WorksheetViewModel(Excel.Worksheet worksheet)
      {
         this.worksheet = worksheet;
         this.name = worksheet.Name;
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

      public string ImagePath
      {
         get
         {
            return "Treeview/Worksheet.bmp";
         }
      }

      internal Excel.Worksheet Worksheet
      {
         get
         {
            return this.worksheet;
         }
      }

      protected void OnPropertyChanged(string propertyName)
      {
         var handler = this.PropertyChanged;
         if (handler != null)
         {
            handler(this, new PropertyChangedEventArgs(propertyName));
         }
      }

      internal void UpdateDisplayProperties()
      {
         this.Name = this.worksheet.Name;
      }
   }
}