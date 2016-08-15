//-----------------------------------------------------------------------
// <copyright file="WorksheetsComparer.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane.Comparison
{
   using System;
   using System.Collections.Generic;
   using System.Windows.Forms;
   using Excel = Microsoft.Office.Interop.Excel;

   /// <summary>
   /// The WorksheetsComparer, given 2 Excel worksheets,
   /// is responsible for producing a list of the differences
   /// between the 2. A difference is either a different value,
   /// or a different formula.
   /// </summary>
   public class WorksheetsComparer
   {
      public static List<Difference> FindDifferences(Excel.Worksheet firstSheet, Excel.Worksheet secondSheet)
      {
         var differences = new List<Difference>();

         try
         {
            var lastCellFirst = GetLastCell(firstSheet);
            var lastCellSecond = GetLastCell(secondSheet);

            var rows = Math.Max(lastCellFirst.Row, lastCellSecond.Row);
            var columns = Math.Max(lastCellFirst.Column, lastCellSecond.Column);

            var firstValues = ReadValues(firstSheet, rows, columns);
            var firstFormulas = ReadFormulas(firstSheet, rows, columns);
            var secondValues = ReadValues(secondSheet, rows, columns);
            var secondFormulas = ReadFormulas(secondSheet, rows, columns);

            for (int row = 1; row <= rows; row++)
            {
               for (int column = 1; column <= columns; column++)
               {
                  var firstValue = ConvertToString(firstValues[row, column]);
                  var secondValue = ConvertToString(secondValues[row, column]);
                  var firstFormula = ConvertToString(firstFormulas[row, column]);
                  var secondFormula = ConvertToString(secondFormulas[row, column]);

                  if (firstValue != secondValue || firstFormula != secondFormula)
                  {
                     var difference = new Difference();
                     difference.Row = row;
                     difference.Column = column;
                     difference.OriginalValue = firstValue;
                     difference.OtherValue = secondValue;
                     difference.OriginalFormula = firstFormula;
                     difference.OtherFormula = secondFormula;
                     differences.Add(difference);
                  }
               }
            }
         }
         catch
         {
            var message = string.Format("Failed to read and compare {0} and {1}", firstSheet.Name, secondSheet.Name);
            MessageBox.Show(message);
            differences = new List<Difference>();
         }

         return differences;
      }

      private static Excel.Range GetLastCell(Excel.Worksheet worksheet)
      {
         var lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
         return lastCell;
      }

      private static string ConvertToString(object content)
      {
         if (content == null)
         {
            return string.Empty;
         }

         return Convert.ToString(content);
      }

      private static object[,] ReadContents(Excel.Worksheet sheet, Func<Excel.Range, object> reader, int lastRow, int lastColumn)
      {
         object[,] cellContents;
         var firstCell = sheet.get_Range("A1", Type.Missing);
         var lastCell = (Excel.Range)sheet.Cells[lastRow, lastColumn];

         if (lastRow == 1 && lastColumn == 1)
         {
            cellContents = new object[2, 2];
            cellContents[1, 1] = reader(firstCell);
         }
         else
         {
            Excel.Range worksheetCells = sheet.get_Range(firstCell, lastCell);
            cellContents = reader(worksheetCells) as object[,];
         }

         return cellContents;
      }

      private static object[,] ReadValues(Excel.Worksheet sheet, int lastRow, int lastColumn)
      {
         var reader = new Func<Excel.Range, object>(r => r.Value2);
         object[,] cellValues = ReadContents(sheet, reader, lastRow, lastColumn);
         return cellValues;
      }

      private static object[,] ReadFormulas(Excel.Worksheet sheet, int lastRow, int lastColumn)
      {
         var reader = new Func<Excel.Range, object>(r => r.Formula);
         object[,] cellFormulas = ReadContents(sheet, reader, lastRow, lastColumn);
         return cellFormulas;
      }
   }
}
