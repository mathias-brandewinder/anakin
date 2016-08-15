//-----------------------------------------------------------------------
// <copyright file="AnakinViewModel.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane
{
   using ClearLines.Anakin.TaskPane.TreeView;
   using Excel = Microsoft.Office.Interop.Excel;

   /// <summary>
   /// AnakinViewModel is the intermediary between
   /// the add-in functionality, and the user interface.
   /// It is bound to the AnakinView control, hosted
   /// in the TaskPaneView.
   /// </summary>
   public class AnakinViewModel
   {
      private ExcelViewModel excelViewModel;
      private Excel.Application excel;

      internal AnakinViewModel(Excel.Application excel)
      {
         this.excel = excel;
      }

      public ExcelViewModel ExcelViewModel
      {
         get
         {
            if (this.excelViewModel == null)
            {
               this.excelViewModel = new ExcelViewModel(this.excel);
            }

            return this.excelViewModel;
         }
      }

      internal Excel.Worksheet SelectedWorksheet
      {
         get;
         set;
      }
   }
}
