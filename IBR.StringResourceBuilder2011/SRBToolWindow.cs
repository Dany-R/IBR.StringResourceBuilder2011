
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;



namespace IBR.StringResourceBuilder2011
{
  /// <summary>
  /// This class implements the tool window SRBToolWindow exposed by this package and hosts a user control.
  ///
  /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane, 
  /// usually implemented by the package implementer.
  ///
  /// This class derives from the ToolWindowPane class provided from the MPF in order to use its 
  /// implementation of the IVsUIElementPane interface.
  /// </summary>
  [Guid("95647f78-65e5-4a37-874a-3847c4b24865")]
  public class SRBToolWindow : SRBToolWindowBase
  {
    /// <summary>
    /// Standard constructor for the tool window.
    /// </summary>
    public SRBToolWindow()
    {
      Trace.WriteLine($"Entering constructor for: {this.ToString()}");

      this.Caption = IBR.StringResourceBuilder2011.Properties.Resources.ToolWindowTitle;

      // Set the image that will appear on the tab of the window frame when docked with another window.
      // The resource ID correspond to the one defined in the ResX file while the Index is the offset
      // in the bitmap strip. Each image in the strip being 16x16.
      this.BitmapResourceID = 301;
      this.BitmapIndex      = 0;

      base.Content = new SRBControl();
    }
  } //class
} //namespace
