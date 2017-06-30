using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using EnvDTE;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;



namespace IBR.StringResourceBuilder2011
{
  //these attributes have been copied from base class (IBRStringResourceBuilder2011PackageBase) so that VS can use them
  [InstalledProductRegistration("#110", "#112", "1.6", IconResourceID = 400)] // Info on this package for Help/About
  [Guid(GuidList.guidIBRStringResourceBuilder2011PkgString)]
  public class IBRStringResourceBuilder2011Package : IBRStringResourceBuilder2011PackageBase
  {
    #region Constructor

    static IBRStringResourceBuilder2011Package()
    {
      //if (m_Dte == null)
      //  m_Dte = (DTE)GetGlobalService(typeof(DTE));
    }

    //public IBRStringResourceBuilder2011Package()
    //  : base()
    //{
    //}

    #endregion //Constructor -----------------------------------------------------------------------

    #region Fields

    //private static DTE m_Dte;
    private SRBToolWindow m_Window;
    private SRBControl m_Control;

    #endregion //Fields ----------------------------------------------------------------------------

    #region Handlers for Button: StringResourceBuilder

    protected override void StringResourceBuilderExecuteHandler(object sender, EventArgs e)
    {
      base.StringResourceBuilderExecuteHandler(sender, e);

      //SRB must be set to Transient=true in vspackage designer (no automatic window open on IDE start)
      if (m_Window == null)
      {
        m_Window = FindToolWindow(typeof(SRBToolWindow), 0, false) as SRBToolWindow;
        if (m_Window != null)
          m_Control = m_Window.Content as SRBControl;
      } //if
    }

    protected override void StringResourceBuilderChangeHandler(object sender, EventArgs e)
    {
      base.StringResourceBuilderChangeHandler(sender, e);
    }

    protected override void StringResourceBuilderQueryStatusHandler(object sender, EventArgs e)
    {
      base.StringResourceBuilderQueryStatusHandler(sender, e);
    }

    #endregion

    #region Handlers for Button: Rescan

    protected override void RescanExecuteHandler(object sender, EventArgs e)
    {
      //base.RescanExecuteHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if (m_Control == null)
        return;

      m_Control.DoRescan();
    }

    protected override void RescanChangeHandler(object sender, EventArgs e)
    {
      //base.RescanChangeHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;
    }

    protected override void RescanQueryStatusHandler(object sender, EventArgs e)
    {
      //Debug.Print(">>> RescanQueryStatusHandler()");

      //base.RescanQueryStatusHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if ((m_Control != null) && (m_Control.RescanButton == null))
        m_Control.RescanButton = command;

      command.Enabled = (m_Control == null) ? false : m_Control.GetRescanEnabled();
    }

    #endregion

    #region Handlers for Button: First

    protected override void FirstExecuteHandler(object sender, EventArgs e)
    {
      //base.FirstExecuteHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if (m_Control == null)
        return;

      m_Control.DoGotoFirst();
    }

    protected override void FirstChangeHandler(object sender, EventArgs e)
    {
      //base.FirstChangeHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;
    }

    protected override void FirstQueryStatusHandler(object sender, EventArgs e)
    {
      //Debug.Print(">>> FirstQueryStatusHandler()");

      //base.FirstQueryStatusHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if ((m_Control != null) && (m_Control.FirstButton == null))
        m_Control.FirstButton = command;

      command.Enabled = (m_Control == null) ? false : m_Control.GetFirstEnabled();
    }

    #endregion

    #region Handlers for Button: Previous

    protected override void PreviousExecuteHandler(object sender, EventArgs e)
    {
      OleMenuCommand command = sender as OleMenuCommand;

      //base.PreviousExecuteHandler(sender, e);

      if (command == null)
        return;

      if (m_Control == null)
        return;

      m_Control.DoGotoPrevious();
    }

    protected override void PreviousChangeHandler(object sender, EventArgs e)
    {
      //base.PreviousChangeHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;
    }

    protected override void PreviousQueryStatusHandler(object sender, EventArgs e)
    {
      //Debug.Print(">>> PreviousQueryStatusHandler()");

      //base.PreviousQueryStatusHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if ((m_Control != null) && (m_Control.PreviousButton == null))
        m_Control.PreviousButton = command;

      command.Enabled = (m_Control == null) ? false : m_Control.GetPreviousEnabled();
    }

    #endregion

    #region Handlers for Button: Next

    protected override void NextExecuteHandler(object sender, EventArgs e)
    {
      OleMenuCommand command = sender as OleMenuCommand;

      //base.NextExecuteHandler(sender, e);

      if (command == null)
        return;

      if (m_Control == null)
        return;

      m_Control.DoGotoNext();
    }

    protected override void NextChangeHandler(object sender, EventArgs e)
    {
      //base.NextChangeHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;
    }

    protected override void NextQueryStatusHandler(object sender, EventArgs e)
    {
      //Debug.Print(">>> NextQueryStatusHandler()");

      //base.NextQueryStatusHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if ((m_Control != null) && (m_Control.NextButton == null))
        m_Control.NextButton = command;

      command.Enabled = (m_Control == null) ? false : m_Control.GetNextEnabled();
    }

    #endregion

    #region Handlers for Button: Last

    protected override void LastExecuteHandler(object sender, EventArgs e)
    {
      //base.LastExecuteHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if (m_Control == null)
        return;

      m_Control.DoGotoLast();
    }

    protected override void LastChangeHandler(object sender, EventArgs e)
    {
      //base.LastChangeHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;
    }

    protected override void LastQueryStatusHandler(object sender, EventArgs e)
    {
      //Debug.Print(">>> LastQueryStatusHandler()");

      //base.LastQueryStatusHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if ((m_Control != null) && (m_Control.LastButton == null))
        m_Control.LastButton = command;

      command.Enabled = (m_Control == null) ? false : m_Control.GetLastEnabled();
    }

    #endregion

    #region Handlers for Button: Make

    protected override void MakeExecuteHandler(object sender, EventArgs e)
    {
      //base.MakeExecuteHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if (m_Control == null)
        return;

      m_Control.DoMake();
    }

    protected override void MakeChangeHandler(object sender, EventArgs e)
    {
      //base.MakeChangeHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;
    }

    protected override void MakeQueryStatusHandler(object sender, EventArgs e)
    {
      //Debug.Print(">>> MakeQueryStatusHandler()");

      //base.MakeQueryStatusHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if ((m_Control != null) && (m_Control.MakeButton == null))
        m_Control.MakeButton = command;

      command.Enabled = (m_Control == null) ? false : m_Control.GetMakeEnabled();
    }

    #endregion

    #region Handlers for Button: Settings

    protected override void SettingsExecuteHandler(object sender, EventArgs e)
    {
      //base.SettingsExecuteHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if (m_Control == null)
        return;

      m_Control.DoSettings();
    }

    protected override void SettingsChangeHandler(object sender, EventArgs e)
    {
      //base.SettingsChangeHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;
    }

    protected override void SettingsQueryStatusHandler(object sender, EventArgs e)
    {
      //Debug.Print(">>> SettingsQueryStatusHandler()");

      //base.SettingsQueryStatusHandler(sender, e);

      OleMenuCommand command = sender as OleMenuCommand;

      if (command == null)
        return;

      if ((m_Control != null) && (m_Control.SettingsButton == null))
        m_Control.SettingsButton = command;

      command.Enabled = (m_Control == null) ? false : m_Control.GetSettingsEnabled();
    }

    #endregion
  } //class
} //namespace
