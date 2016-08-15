//-----------------------------------------------------------------------
// <copyright file="ExcelViewModel.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane.TreeView
{
   using System.Collections.ObjectModel;
   using Excel = Microsoft.Office.Interop.Excel;

   /// <summary>
   /// ExcelViewModel is the root element of the
   /// TreeView displayed in the AnakinView control.
   /// It holds a collection of WorkbookViewModel,
   /// which represent the workbooks currently open.
   /// It is responsible for observing events in Excel,
   /// and propagating updates through the workbooks.
   /// Because of the shortcomings of the events Excel
   /// exposes, updates are triggered when the active
   /// item is changed, refreshing the tree.
   /// </summary>
   public class ExcelViewModel
   {
      private ObservableCollection<WorkbookViewModel> workbookViewModels;
      private Excel.Application excel;

      internal ExcelViewModel(Excel.Application excel)
      {
         this.workbookViewModels = new ObservableCollection<WorkbookViewModel>();
         this.excel = excel;
         ((Excel.AppEvents_Event)excel).NewWorkbook += this.AddWorkbook;
         excel.WorkbookOpen += this.AddWorkbook;
         excel.WorkbookActivate += this.ActiveWorkbookChanged;
         excel.SheetActivate += this.ActiveSheetChanged;
         var workbooks = excel.Workbooks;
         foreach (var workbook in workbooks)
         {
            var book = workbook as Excel.Workbook;
            if (book != null)
            {
               var workbookViewModel = new WorkbookViewModel(book);
               this.workbookViewModels.Add(workbookViewModel);
            }
         }
      }

      public ObservableCollection<WorkbookViewModel> Workbooks
      {
         get
         {
            return this.workbookViewModels;
         }
      }

      private void ActiveSheetChanged(object activatedSheet)
      {
         this.UpdateWorkbooks();
      }

      private void ActiveWorkbookChanged(Excel.Workbook activatedWorkbook)
      {
         this.UpdateWorkbooks();
      }

      private void UpdateWorkbooks()
      {
         var workbooks = this.excel.Workbooks;
         foreach (var workbookViewModel in this.workbookViewModels)
         {
            var workbookIsOpen = false;
            foreach (var workbook in workbooks)
            {
               if (workbookViewModel.Workbook == workbook)
               {
                  workbookIsOpen = true;
                  break;
               }
            }

            if (workbookIsOpen == false)
            {
               this.workbookViewModels.Remove(workbookViewModel);
            }
            else
            {
               workbookViewModel.UpdateDisplayProperties();
               this.UpdateWorksheets(workbookViewModel);
            }
         }
      }

      private void UpdateWorksheets(WorkbookViewModel workbookViewModel)
      {
         var workbook = workbookViewModel.Workbook;
         var worksheets = workbook.Worksheets;
         foreach (var worksheetViewModel in workbookViewModel.Worksheets)
         {
            var worksheetIsOpen = false;
            foreach (var sheet in worksheets)
            {
               var worksheet = sheet as Excel.Worksheet;
               if (worksheet != null)
               {
                  if (worksheet == worksheetViewModel.Worksheet)
                  {
                     worksheetIsOpen = true;
                     break;
                  }
               }
            }

            if (worksheetIsOpen == false)
            {
               workbookViewModel.Worksheets.Remove(worksheetViewModel);
            }
            else
            {
               worksheetViewModel.UpdateDisplayProperties();
            }
         }
      }

      private void AddWorkbook(Excel.Workbook newWorkbook)
      {
         var workbookViewModel = new WorkbookViewModel(newWorkbook);
         this.workbookViewModels.Add(workbookViewModel);
      }
   }
}