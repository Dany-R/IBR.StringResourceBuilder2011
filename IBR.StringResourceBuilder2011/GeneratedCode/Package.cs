using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace IBR.StringResourceBuilder2011
{
  /// <summary>
  /// This is the class that implements the package exposed by this assembly.
  /// </summary>
  /// <remarks>
  /// <para>
  /// The minimum requirement for a class to be considered a valid package for Visual Studio
  /// is to implement the IVsPackage interface and register itself with the shell.
  /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
  /// to do it: it derives from the Package class that provides the implementation of the
  /// IVsPackage interface and uses the registration attributes defined in the framework to
  /// register itself and its components with the shell. These attributes tell the pkgdef creation
  /// utility what data to put into .pkgdef file.
  /// </para>
  /// <para>
  /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
  /// </para>
  /// </remarks>
  [PackageRegistration(UseManagedResourcesOnly = true)]
  [InstalledProductRegistration("#110", "#112", "1.6", IconResourceID = 400)] // Info on this package for Help/About
  [Guid(GuidList.guidIBRStringResourceBuilder2011PkgString)]
  [ProvideMenuResource("Menus.ctmenu", 1)]  // This attribute is needed to let the shell know that this package exposes some menus.
  [ProvideToolWindow(typeof(SRBToolWindow), Orientation=ToolWindowOrientation.Right, Style=VsDockStyle.Float, MultiInstances = false, Transient = true, PositionX = 100 , PositionY = 100 , Width = 600 , Height = 300 )]
  public abstract class IBRStringResourceBuilder2011PackageBase : Package
  {
    /// <summary>
    /// Default constructor of the package.
    /// Inside this method you can place any initialization code that does not require
    /// any Visual Studio service because at this point the package object is created but
    /// not sited yet inside Visual Studio environment. The place to do all the other
    /// initialization is the Initialize method.
    /// </summary>
    public IBRStringResourceBuilder2011PackageBase()
    {
      Trace.WriteLine($"Entering constructor for: {this.ToString()}");
    }

    /////////////////////////////////////////////////////////////////////////////
    // Overridden Package Implementation

    #region Package Members

    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited, so this is the place
    /// where you can put all the initialization code that rely on services provided by VisualStudio.
    /// </summary>
    protected override void Initialize()
    {
      Trace.WriteLine ($"Entering Initialize() of: {this.ToString()}");
      base.Initialize();

      // Add our command handlers for menu (commands must exist in the .vsct file)
      OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      if ( null != mcs )
      {
        CommandID commandId;
        OleMenuCommand menuItem;

        // Create the command for button StringResourceBuilder
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.StringResourceBuilder);
        menuItem = new OleMenuCommand(StringResourceBuilderExecuteHandler, StringResourceBuilderChangeHandler, StringResourceBuilderQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);

        // Create the command for button Rescan
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.Rescan);
        menuItem = new OleMenuCommand(RescanExecuteHandler, RescanChangeHandler, RescanQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);

        // Create the command for button First
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.First);
        menuItem = new OleMenuCommand(FirstExecuteHandler, FirstChangeHandler, FirstQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);

        // Create the command for button Previous
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.Previous);
        menuItem = new OleMenuCommand(PreviousExecuteHandler, PreviousChangeHandler, PreviousQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);

        // Create the command for button Next
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.Next);
        menuItem = new OleMenuCommand(NextExecuteHandler, NextChangeHandler, NextQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);

        // Create the command for button Last
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.Last);
        menuItem = new OleMenuCommand(LastExecuteHandler, LastChangeHandler, LastQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);

        // Create the command for button Make
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.Make);
        menuItem = new OleMenuCommand(MakeExecuteHandler, MakeChangeHandler, MakeQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);

        // Create the command for button Settings
        commandId = new CommandID(GuidList.guidIBRStringResourceBuilder2011CmdSet, (int)PkgCmdIDList.Settings);
        menuItem = new OleMenuCommand(SettingsExecuteHandler, SettingsChangeHandler, SettingsQueryStatusHandler, commandId);
        mcs.AddCommand(menuItem);
      } //if
    }

    #endregion

    #region Handlers for Button: StringResourceBuilder

    protected virtual void StringResourceBuilderExecuteHandler(object sender, EventArgs e)
    {
      ShowToolWindowSRB(sender, e);
    }

    protected virtual void StringResourceBuilderChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void StringResourceBuilderQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    #region Handlers for Button: Rescan

    protected virtual void RescanExecuteHandler(object sender, EventArgs e)
    {
      ShowMessage("Rescan clicked!");
    }

    protected virtual void RescanChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void RescanQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    #region Handlers for Button: First

    protected virtual void FirstExecuteHandler(object sender, EventArgs e)
    {
      ShowMessage("First clicked!");
    }

    protected virtual void FirstChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void FirstQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    #region Handlers for Button: Previous

    protected virtual void PreviousExecuteHandler(object sender, EventArgs e)
    {
      ShowMessage("Previous clicked!");
    }

    protected virtual void PreviousChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void PreviousQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    #region Handlers for Button: Next

    protected virtual void NextExecuteHandler(object sender, EventArgs e)
    {
      ShowMessage("Next clicked!");
    }

    protected virtual void NextChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void NextQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    #region Handlers for Button: Last

    protected virtual void LastExecuteHandler(object sender, EventArgs e)
    {
      ShowMessage("Last clicked!");
    }

    protected virtual void LastChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void LastQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    #region Handlers for Button: Make

    protected virtual void MakeExecuteHandler(object sender, EventArgs e)
    {
      ShowMessage("Make clicked!");
    }

    protected virtual void MakeChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void MakeQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    #region Handlers for Button: Settings

    protected virtual void SettingsExecuteHandler(object sender, EventArgs e)
    {
      ShowMessage("Settings clicked!");
    }

    protected virtual void SettingsChangeHandler(object sender, EventArgs e)
    {
    }

    protected virtual void SettingsQueryStatusHandler(object sender, EventArgs e)
    {
    }

    #endregion

    /// <summary>
    /// This function is called when the user clicks the menu item that shows the
    /// tool window. See the Initialize method to see how the menu item is associated to
    /// this function using the OleMenuCommandService service and the MenuCommand class.
    /// </summary>
    private void ShowToolWindowSRB(object sender, EventArgs e)
    {
      // Get the instance number 0 of this tool window. This window is single instance so this instance
      // is actually the only one.
      // The last flag is set to true so that if the tool window does not exists it will be created.
      ToolWindowPane window = FindToolWindow(typeof(SRBToolWindow), 0, true);
      if ((null == window) || (null == window.Frame))
        throw new NotSupportedException("Can not create Toolwindow: SRB");

      IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
      Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
    }

    /// <summary>
    /// This function is the callback used to execute a command when the a menu item is clicked.
    /// See the Initialize method to see how the menu item is associated to this function using
    /// the OleMenuCommandService service and the MenuCommand class.
    /// </summary>
    protected void ShowMessage(string message)
    {
      // Show a Message Box to prove we were here
      IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
      Guid clsid = Guid.Empty;
      int result;
      Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
             0,
             ref clsid,
             "IBR.StringResourceBuilder2011",
             string.Format(CultureInfo.CurrentCulture, message, this.ToString()),
             string.Empty,
             0,
             OLEMSGBUTTON.OLEMSGBUTTON_OK,
             OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
             OLEMSGICON.OLEMSGICON_INFO,
             0,    // false
             out result));
    }
  } //class
} //namespace
