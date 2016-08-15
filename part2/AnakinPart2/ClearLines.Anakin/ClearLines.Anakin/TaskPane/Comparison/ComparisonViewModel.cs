//-----------------------------------------------------------------------
// <copyright file="ComparisonViewModel.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane.Comparison
{
   using System;
   using System.Collections.Generic;
   using System.ComponentModel;
   using System.Windows.Input;
   using Excel = Microsoft.Office.Interop.Excel;

   public class ComparisonViewModel : INotifyPropertyChanged
   {
      private readonly Excel.Application excel;
      private readonly List<Difference> differences;
      private Difference selectedDifference;
      private ICommand goToNextDifference;
      private ICommand goToPreviousDifference;

      public ComparisonViewModel(Excel.Application excel)
      {
         this.excel = excel;
         this.differences = new List<Difference>();
      }

      public event PropertyChangedEventHandler PropertyChanged;

      public Difference SelectedDifference
      {
         get
         {
            return this.selectedDifference;
         }

         set
         {
            if (this.selectedDifference != value)
            {
               this.selectedDifference = value;
               this.NavigateToDifference(this.SelectedDifference);
               this.OnPropertyChanged("SelectedDifference");
            }
         }
      }

      public ICommand GoToNextDifference
      {
         get
         {
            if (this.goToNextDifference == null)
            {
               this.goToNextDifference = new RelayCommand(GoToNextDifferenceExecute, CanGoToNextDifference);
            }

            return this.goToNextDifference;
         }
      }

      public ICommand GoToPreviousDifference
      {
         get
         {
            if (this.goToPreviousDifference == null)
            {
               this.goToPreviousDifference = new RelayCommand(GoToPreviousDifferenceExecute, CanGoToPreviousDifference);
            }

            return this.goToPreviousDifference;
         }
      }

      internal void SetDifferences(List<Difference> newDifferences)
      {
         this.differences.Clear();
         if (newDifferences != null)
         {
            this.differences.AddRange(newDifferences);
         }

         if (this.differences.Count > 0)
         {
            this.SelectedDifference = this.differences[0];
         }
         else
         {
            this.SelectedDifference = null;
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

      private void GoToNextDifferenceExecute(object target)
      {
         var currentIndex = this.differences.IndexOf(SelectedDifference);
         currentIndex++;
         this.SelectedDifference = this.differences[currentIndex];
      }

      private bool CanGoToNextDifference(object arg)
      {
         if (this.DifferencesAreNullOrEmpty())
         {
            return false;
         }

         return (this.differences.IndexOf(SelectedDifference) < this.differences.Count - 1);
      }

      private void GoToPreviousDifferenceExecute(object target)
      {
         var currentIndex = this.differences.IndexOf(SelectedDifference);
         currentIndex--;
         this.SelectedDifference = this.differences[currentIndex];
      }

      private bool CanGoToPreviousDifference(object arg)
      {
         if (this.DifferencesAreNullOrEmpty())
         {
            return false;
         }

         return (this.differences.IndexOf(SelectedDifference) > 0);
      }

      private void NavigateToDifference(Difference difference)
      {
         if (difference == null)
         {
            return;
         }

         var row = difference.Row;
         var column = difference.Column;
         var activeSheet = (Excel.Worksheet)this.excel.ActiveSheet;
         var differenceLocation = activeSheet.Cells[row, column];
         this.excel.Goto(differenceLocation, Type.Missing);
      }

      private bool DifferencesAreNullOrEmpty()
      {
         if (this.differences == null)
         {
            return true;
         }

         if (this.differences.Count == 0)
         {
            return true;
         }

         return false;
      }
   }
}
