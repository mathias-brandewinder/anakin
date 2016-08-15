//-----------------------------------------------------------------------
// <copyright file="HasDifferenceConverter.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane.Comparison
{
   using System;
   using System.Globalization;
   using System.Windows.Data;
   using System.Windows.Media;

   /// <summary>
   /// HasDifferenceConverter is a WPF converter;
   /// It takes in a boolean which indicates whether
   /// there is a difference between two entities,
   /// and returns a Brush, used to 
   /// indicate visually that there is a difference.
   /// </summary>
   [ValueConversion(typeof(bool), typeof(Brushes))]
   public class HasDifferenceConverter : IValueConverter
   {
      public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
      {
         var hasDifference = (bool)value;
         if (hasDifference)
         {
            return Brushes.Orange;
         }

         return Brushes.GhostWhite;
      }

      public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
      {
         return null;
      }
   }
}
