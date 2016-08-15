//-----------------------------------------------------------------------
// <copyright file="RelayCommand.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

// This class has been lifted from the MVVMFoundation library
// written by Josh Smith, hosted on CodePlex at 
// http://mvvmfoundation.codeplex.com/
// A few minor cosmetic changes have been made.
namespace ClearLines.Anakin.TaskPane
{
   using System;
   using System.Diagnostics;
   using System.Windows.Input;

   public class RelayCommand : ICommand
   {
      private readonly Action<object> execute;
      private readonly Predicate<object> canExecute;

      public RelayCommand(Action<object> execute)
         : this(execute, null)
      {
      }

      public RelayCommand(Action<object> execute, Predicate<object> canExecute)
      {
         if (execute == null)
         {
            throw new ArgumentNullException("execute");
         }

         this.execute = execute;
         this.canExecute = canExecute;
      }

      [DebuggerStepThrough]
      public bool CanExecute(object parameter)
      {
         return canExecute == null ? true : canExecute(parameter);
      }

      public event EventHandler CanExecuteChanged
      {
         add
         {
            CommandManager.RequerySuggested += value;
         }

         remove
         {
            CommandManager.RequerySuggested -= value;
         }
      }

      public void Execute(object parameter)
      {
         execute(parameter);
      }
   }
}
