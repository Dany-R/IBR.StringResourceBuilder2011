using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel.Design;

namespace IBR.StringResourceBuilder2011
{

  /// <summary>
  /// This class implements the tool window SRBToolWindowBase exposed by this package and hosts a user control.
  ///
  /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
  /// usually implemented by the package implementer.
  ///
  /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
  /// implementation of the IVsUIElementPane interface.
  /// </summary>
  [Guid("95647f78-65e5-4a37-874a-3847c4b24865")]
  public class SRBToolWindowBase : ToolWindowPane
  {
    /// <summary>
    /// Standard constructor for the tool window.
    /// </summary>
    public SRBToolWindowBase()
      : base(null)
    {
      this.Caption = "SRB";
      this.ToolBar = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.SRBToolbarMenu);
    }
  } //class
} //namespace
