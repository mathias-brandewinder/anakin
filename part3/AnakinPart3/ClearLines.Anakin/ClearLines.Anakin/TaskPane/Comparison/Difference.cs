//-----------------------------------------------------------------------
// <copyright file="Difference.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane.Comparison
{
   public class Difference
   {
      public string OriginalValue
      {
         get;
         set;
      }

      public string OtherValue
      {
         get;
         set;
      }

      public string OriginalFormula
      {
         get;
         set;
      }

      public string OtherFormula
      {
         get;
         set;
      }

      public bool AreValuesDifferent
      {
         get
         {
            return this.OriginalValue != this.OtherValue;
         }
      }

      public bool AreFormulasDifferent
      {
         get
         {
            return this.OriginalFormula != this.OtherFormula;
         }
      }

      public int Row
      {
         get;
         set;
      }

      public int Column
      {
         get;
         set;
      }

      public string Location
      {
         get
         {
            return "R" + this.Row + "C" + this.Column;
         }
      }
   }
}
