//-----------------------------------------------------------------------
// <copyright file="ThisAddIn.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin
{
   using ClearLines.Anakin.TaskPane;
   using Microsoft.Office.Tools;

   /// <summary>
   /// ThisAddIn is the entry point that gets instantiated
   /// when the add-in loads. The Start-up methods wires
   /// together the task pane control, the AnakinView control
   /// and the AnakinViewModel. 
   /// The TaskPane is exposed internally so that the Ribbon
   /// can access it and show/hide it.
   /// </summary>
   public partial class ThisAddIn
   {
      private CustomTaskPane taskPane;

      internal CustomTaskPane TaskPane
      {
         get
         {
            return this.taskPane;
         }
      }

      private void ThisAddIn_Startup(object sender, System.EventArgs e)
      {
         var taskPaneView = new TaskPaneView();
         this.taskPane = this.CustomTaskPanes.Add(taskPaneView, "Anakin");
         this.taskPane.Visible = false;

         var excel = this.Application;
         var anakinViewModel = new AnakinViewModel(excel);
         var anakinView = taskPaneView.AnakinView;
         anakinView.DataContext = anakinViewModel;
      }

      private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
      {
      }

      #region VSTO generated code

      /// <summary>
      /// Required method for Designer support - do not modify
      /// the contents of this method with the code editor.
      /// </summary>
      private void InternalStartup()
      {
         this.Startup += new System.EventHandler(ThisAddIn_Startup);
         this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
      }

      #endregion
   }
}
