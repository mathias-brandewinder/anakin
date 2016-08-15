//-----------------------------------------------------------------------
// <copyright file="TaskPaneView.cs" company="Clear Lines Consulting, LLC">
//     Copyright (c) Clear Lines Consulting, LLC. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace ClearLines.Anakin.TaskPane
{
   using System.Windows.Forms;

   public partial class TaskPaneView : UserControl
   {
      public TaskPaneView()
      {
         InitializeComponent();
      }

      internal AnakinView AnakinView
      {
         get
         {
            return this.anakinView;
         }
      }
   }
}
