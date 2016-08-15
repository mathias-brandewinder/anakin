//-----------------------------------------------------------------------
// <copyright file="AnakinViewModel.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane
{
   using System.Windows.Input;
   using ClearLines.Anakin.TaskPane.Comparison;
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
      private readonly ComparisonViewModel comparisonViewModel;
      private readonly Excel.Application excel;
      private ExcelViewModel excelViewModel;
      private ICommand generateComparison;

      internal AnakinViewModel(Excel.Application excel)
      {
         this.excel = excel;
         this.comparisonViewModel = new ComparisonViewModel(excel);
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

      public ICommand GenerateComparison
      {
         get
         {
            if (this.generateComparison == null)
            {
               this.generateComparison = new RelayCommand(this.GenerateComparisonExecute, this.CanGenerateComparison);
            }

            return this.generateComparison;
         }
      }

      internal ComparisonViewModel ComparisonViewModel
      {
         get
         {
            return this.comparisonViewModel;
         }
      }

      internal Excel.Worksheet SelectedWorksheet
      {
         get;
         set;
      }

      private void GenerateComparisonExecute(object target)
      {
         var currentSheet = this.excel.ActiveSheet as Excel.Worksheet;
         var selectedSheet = this.SelectedWorksheet;

         var differences = WorksheetsComparer.FindDifferences(currentSheet, selectedSheet);
         this.comparisonViewModel.SetDifferences(differences);
      }

      private bool CanGenerateComparison(object target)
      {
         return this.SelectedWorksheet != null;
      }
   }
}
