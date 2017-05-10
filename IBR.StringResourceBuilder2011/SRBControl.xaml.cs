using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using EnvDTE;
using EnvDTE80;

using IBR.StringResourceBuilder2011.Modules;



namespace IBR.StringResourceBuilder2011
{
  /// <summary>
  /// Interaction logic for SRBControl.xaml
  /// </summary>
  public partial class SRBControl : UserControl
  {
    #region Constructor

    public SRBControl()
    {
      Trace.WriteLine("Entering SRBControl()");

      InitializeComponent();

      this.progressBar1.Visibility = Visibility.Hidden;

      m_Dte2 = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE)) as DTE2;

      m_StringResourceBuilder = new StringResourceBuilder(m_Dte2);

      m_StringResourceBuilder.InitProgress        = InitProgress;
      m_StringResourceBuilder.HideProgress        = HideProgress;
      m_StringResourceBuilder.SetProgress         = SetProgress;
      m_StringResourceBuilder.ClearGrid           = ClearGrid;
      m_StringResourceBuilder.SetGridItemsSource  = SetGridItemsSource;
      m_StringResourceBuilder.RefreshGrid         = RefreshGrid;
      m_StringResourceBuilder.SelectCell          = SelectCell;
      m_StringResourceBuilder.GetSelectedItem     = GetSelectedItem;
      m_StringResourceBuilder.GetSelectedColIndex = GetSelectedColIndex;
      m_StringResourceBuilder.GetSelectedRowIndex = GetSelectedRowIndex;
    }

    #endregion //Constructor -----------------------------------------------------------------------

    #region Types
    #endregion //Types -----------------------------------------------------------------------------

    #region Fields

    private DTE2 m_Dte2;

    private bool m_IsVisible;
    //private bool m_IsVisibleChanged;

    private StringResourceBuilder m_StringResourceBuilder;

    private TextEditorEvents m_TextEditorEvents;
    //private EnvDTE80.TextDocumentKeyPressEvents m_TextDocumentKeyPressEvents;
    private int m_FocusedTextDocumentWindowHash;

    private bool m_IsSelectedCellChangedProgrammatically;
    private bool m_IsMakeInProgress;

    private double m_Progress;
    private double m_ProgressStep;

    private int m_LastCurrentColumn = 1;

    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties

    private Microsoft.VisualStudio.Shell.OleMenuCommand m_RescanButton;
    public Microsoft.VisualStudio.Shell.OleMenuCommand RescanButton
    {
      get { return (m_RescanButton); }
      set { m_RescanButton = value; }
    }

    private Microsoft.VisualStudio.Shell.OleMenuCommand m_FirstButton;
    public Microsoft.VisualStudio.Shell.OleMenuCommand FirstButton
    {
      get { return (m_FirstButton); }
      set { m_FirstButton = value; }
    }

    private Microsoft.VisualStudio.Shell.OleMenuCommand m_PreviousButton;
    public Microsoft.VisualStudio.Shell.OleMenuCommand PreviousButton
    {
      get { return (m_PreviousButton); }
      set { m_PreviousButton = value; }
    }

    private Microsoft.VisualStudio.Shell.OleMenuCommand m_NextButton;
    public Microsoft.VisualStudio.Shell.OleMenuCommand NextButton
    {
      get { return (m_NextButton); }
      set { m_NextButton = value; }
    }

    private Microsoft.VisualStudio.Shell.OleMenuCommand m_LastButton;
    public Microsoft.VisualStudio.Shell.OleMenuCommand LastButton
    {
      get { return (m_LastButton); }
      set { m_LastButton = value; }
    }

    private Microsoft.VisualStudio.Shell.OleMenuCommand m_MakeButton;
    public Microsoft.VisualStudio.Shell.OleMenuCommand MakeButton
    {
      get { return (m_MakeButton); }
      set { m_MakeButton = value; }
    }

    private Microsoft.VisualStudio.Shell.OleMenuCommand m_SettingsButton;
    public Microsoft.VisualStudio.Shell.OleMenuCommand SettingsButton
    {
      get { return (m_SettingsButton); }
      set { m_SettingsButton = value; }
    }

    #endregion //Properties ------------------------------------------------------------------------

    #region Events

    private void StringResBuilderWindow_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
      m_IsVisible = (bool)e.NewValue;
      //m_IsVisibleChanged = true;

      if (!m_IsVisible)
      {
        //remove event handlers
        m_Dte2.Events.WindowEvents.WindowActivated -= WindowEvents_WindowActivated;
        m_Dte2.Events.WindowEvents.WindowClosing   -= WindowEvents_WindowClosing;

        Cleanup();
      }
      else
      {
        //assign event handlers (remove first to prevent from doubling them)
        m_Dte2.Events.WindowEvents.WindowActivated -= WindowEvents_WindowActivated;
        m_Dte2.Events.WindowEvents.WindowActivated += WindowEvents_WindowActivated;
        m_Dte2.Events.WindowEvents.WindowClosing   -= WindowEvents_WindowClosing;
        m_Dte2.Events.WindowEvents.WindowClosing   += WindowEvents_WindowClosing;
      } //else
    }

    private void dataGrid1_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
    {
      if (m_IsSelectedCellChangedProgrammatically)
        return;

      ResetCurrentCellToSelection();

      SetEnabled();
    }

    private void dataGrid1_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
      DataGrid grid = (sender as DataGrid);
      DataGridColumn col = grid.CurrentColumn;
      if (col == null)
        return;

      if (!col.IsReadOnly)
        return;

      if (e != null)
        e.Handled = true;

      m_StringResourceBuilder.SelectStringInTextDocument();
    }

    private void dataGrid1_GotFocus(object sender, RoutedEventArgs e)
    {
      if (this.dataGrid1.IsMouseOver)
        return;

      //sadly, the DataGrid sets the focus to the first cell when focus returns so we restore it here
      ResetCurrentCellToSelection();
    }

    #region EnvDTE

    private void WindowEvents_WindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
    {
      //Trace.WriteLine(string.Format("\n>>>> WindowEvents_WindowActivated() - '{0}' <- '{1}' ({2})\n", gotFocus.Caption, lostFocus.Caption, (m_Dte2.ActiveDocument == null) ? "null" : m_Dte2.ActiveDocument.Name));
      Debug.Print("\n>>>> WindowEvents_WindowActivated() - '{0}' <- '{1}' ({2})\n", (gotFocus == null) ? "null" : gotFocus.Caption, (lostFocus == null) ? "null" : lostFocus.Caption, (m_Dte2.ActiveDocument == null) ? "null" : m_Dte2.ActiveDocument.Name);

      if (m_StringResourceBuilder.IsBrowsing)
        return;

      if (!m_IsVisible)
        return;

      if (m_Dte2.ActiveDocument == null)
      {
        if (m_StringResourceBuilder.Window != null)
          WindowEvents_WindowClosing(m_StringResourceBuilder.Window);

        return;
      } //if

      Trace.WriteLine("WindowEvents_WindowActivated()");

      if (m_TextEditorEvents != null)
      {
        m_TextEditorEvents.LineChanged -= m_TextEditorEvents_LineChanged;
        m_TextEditorEvents = null;
      } //if

      //if (m_TextDocumentKeyPressEvents != null)
      //{
      //  m_TextDocumentKeyPressEvents.AfterKeyPress -= m_TextDocumentKeyPressEvents_AfterKeyPress;
      //  m_TextDocumentKeyPressEvents.AfterKeyPress += m_TextDocumentKeyPressEvents_AfterKeyPress;
      //  m_TextDocumentKeyPressEvents = null;
      //} //if

      EnvDTE.Window window = m_Dte2.ActiveDocument.ActiveWindow;

      bool isGotDocument    = (gotFocus != null) && (gotFocus.Document != null),
           isLostDocument   = (lostFocus != null) && (lostFocus.Document != null),
           isGotCode        = isGotDocument && (gotFocus.Caption.EndsWith(".cs") || gotFocus.Caption.EndsWith(".vb")),
           isLostCode       = isLostDocument && (lostFocus.Caption.EndsWith(".cs") || lostFocus.Caption.EndsWith(".vb"));

      if (!isGotDocument /*&& !isLostDocument*/ && (m_StringResourceBuilder.Window != null))
        window = m_StringResourceBuilder.Window;
      if (isGotCode)
        window = gotFocus;
      else if (!isGotDocument && isLostCode)
        window = lostFocus;

      if ((window != null) && (window.Document != null) && (window.Caption.EndsWith(".cs") || window.Caption.EndsWith(".vb")))
      {
        TextDocument txtDoc = window.Document.Object("TextDocument") as TextDocument;
        if (txtDoc != null)
        {
          EnvDTE80.Events2 events = (EnvDTE80.Events2)m_Dte2.Events;

          m_TextEditorEvents = events.TextEditorEvents[txtDoc];
          m_TextEditorEvents.LineChanged -= m_TextEditorEvents_LineChanged;
          m_TextEditorEvents.LineChanged += m_TextEditorEvents_LineChanged;

          //m_TextDocumentKeyPressEvents = events.TextDocumentKeyPressEvents[txtDoc];
          //m_TextDocumentKeyPressEvents.AfterKeyPress -= m_TextDocumentKeyPressEvents_AfterKeyPress;
          //m_TextDocumentKeyPressEvents.AfterKeyPress += m_TextDocumentKeyPressEvents_AfterKeyPress;

          if (m_FocusedTextDocumentWindowHash != window.GetHashCode())
          {
            m_FocusedTextDocumentWindowHash                   = window.GetHashCode();
            m_StringResourceBuilder.FocusedTextDocumentWindow = window;
            this.Dispatcher.BeginInvoke(new Action(m_StringResourceBuilder.DoBrowse));
          } //if
        } //if
      }
      //else if (m_StringResourceBuilder.Window != null)
      //  WindowEvents_WindowClosing(m_StringResourceBuilder.Window);
    }

    private void WindowEvents_WindowClosing(EnvDTE.Window Window)
    {
      if (Window != m_StringResourceBuilder.Window)
        return;

      Trace.WriteLine("WindowEvents_WindowClosing()");

      Cleanup();
    }

    private void m_TextEditorEvents_LineChanged(TextPoint startPoint, TextPoint endPoint, int hint)
    {
      if (m_IsMakeInProgress)
        return;

      vsTextChanged textChangedHint = (vsTextChanged)hint;

      System.Diagnostics.Trace.WriteLine(string.Format("#### m_TextEditorEvents_LineChanged {0};{1} {2};{3} ({4})",
                                                       startPoint.Line, startPoint.LineCharOffset,
                                                       endPoint.Line, endPoint.LineCharOffset,
                                                       textChangedHint.ToString()));

      //if (((textChangedHint & vsTextChanged.vsTextChangedNewline) == 0) &&
      //    ((textChangedHint & vsTextChanged.vsTextChangedMultiLine) == 0) &&
      //    (textChangedHint != 0))
      //  return;

      if (m_IsVisible)
        this.Dispatcher./*Begin*/Invoke(new Action<TextPoint, TextPoint>(m_StringResourceBuilder.DoBrowse), startPoint, endPoint); //[12-07-21 DR]: synchronously due to multiple events for one edit
      else
        m_FocusedTextDocumentWindowHash = 0;
    }

    //private void m_TextDocumentKeyPressEvents_AfterKeyPress(string keyPress, EnvDTE.TextSelection selection, bool inStatementCompletion)
    //{
    //  System.Diagnostics.Trace.WriteLine(string.Format("#### m_TextDocumentKeyPressEvents_AfterKeyPress '{0}' {1} ({2})", keyPress, selection.AnchorPoint.Line, inStatementCompletion));
    //}

    #endregion

    #endregion //Events ----------------------------------------------------------------------------

    #region Private members

    private void SetEnabled()
    {
      int count         = this.dataGrid1.Items.Count,
          selectedIndex = GetSelectedRowIndex();
      bool isEmpty      = (count == 0),
           //isMulti      = (count > 1),
           isFirst      = (selectedIndex == 0),
           isLast       = (selectedIndex == count - 1);

      if (m_FirstButton    != null) m_FirstButton.Enabled    = !isFirst;
      if (m_LastButton     != null) m_LastButton.Enabled     = !isLast;
      if (m_PreviousButton != null) m_PreviousButton.Enabled = !isFirst;
      if (m_NextButton     != null) m_NextButton.Enabled     = !isLast;
      if (m_MakeButton     != null) m_MakeButton.Enabled     = !isEmpty;
    }

    private void Cleanup()
    {
      m_StringResourceBuilder.Window                    = null;
      m_StringResourceBuilder.FocusedTextDocumentWindow = null;
      m_FocusedTextDocumentWindowHash                   = 0;

      if (m_TextEditorEvents != null)
      {
        m_TextEditorEvents.LineChanged -= m_TextEditorEvents_LineChanged;
        m_TextEditorEvents = null;
      } //if

      ClearGrid();

      m_StringResourceBuilder.ClearStringResources();
    }

    #region ProgressBar
    private void InitProgress(double maximum)
    {
      this.progressBar1.Value = 0D;
      this.progressBar1.Maximum = maximum;
      this.progressBar1.Visibility = Visibility.Visible;
      this.progressBar1.Refresh();

      m_Progress = 0D;
      m_ProgressStep = this.progressBar1.Maximum / 10D;
    }

    private void HideProgress()
    {
      this.progressBar1.Visibility = Visibility.Hidden;
      this.progressBar1.Refresh();
    }

    private void SetProgress(double value)
    {
      try
      {
        //System.Diagnostics.Debug.Print("{0}/{1}", value, this.progressBar1.Maximum);
        this.progressBar1.Value = value;

        if (m_Progress + m_ProgressStep <= value)
        {
          m_Progress += m_ProgressStep;
          this.progressBar1.Refresh();
        } //if
      }
      catch
      {
        System.Diagnostics.Debug.Print("###err");
      }
    }

    private void SetProgress(int value)
    {
#if USE_THREAD
      try
      {
        //System.Diagnostics.Debug.Print("{0}/{1}", value, this.progressBar1.Maximum);
        this.progressBar1.Value = (double)value;

        if (m_Progress + m_ProgressStep <= (double)value)
        {
          m_Progress += m_ProgressStep;
          //this.progressBar1.Refresh();
        } //if
      }
      catch
      {
        System.Diagnostics.Debug.Print("###err");
      }
#else
      SetProgress((double)value);
#endif
    }
    #endregion

    #region Grid
    private void ClearGrid()
    {
      if (this.dataGrid1.ItemsSource != null)
      {
        if (this.dataGrid1.CurrentColumn != null)
          m_LastCurrentColumn = this.dataGrid1.CurrentColumn.DisplayIndex;
        else if ((this.dataGrid1.SelectedCells.Count > 0) && (this.dataGrid1.SelectedCells[0].Column != null))
          m_LastCurrentColumn = this.dataGrid1.SelectedCells[0].Column.DisplayIndex;
        else
          m_LastCurrentColumn = 1;
      } //if

      this.dataGrid1.UnselectAllCells();
      this.dataGrid1.ItemsSource = null;
    }

    private void SetGridItemsSource(System.Collections.IEnumerable source)
    {
      this.dataGrid1.ItemsSource = source;
    }

    private void RefreshGrid()
    {
      this.dataGrid1.Items.Refresh();
    }

    private void SelectCell(int rowNo,
                            int colNo)
    {
      if (this.dataGrid1.Items.Count <= rowNo)
        rowNo = this.dataGrid1.Items.Count - 1;

      if (this.dataGrid1.Columns.Count <= colNo)
        colNo = this.dataGrid1.Columns.Count - 1;

      if (rowNo < 0)
        return;

      if (colNo < 0)
        return;

      m_IsSelectedCellChangedProgrammatically = true;

      this.dataGrid1.UnselectAllCells();

      DataGridCellInfo cellInfo = new DataGridCellInfo(this.dataGrid1.Items[rowNo], this.dataGrid1.Columns[colNo]);
      this.dataGrid1.CurrentCell = cellInfo;
      this.dataGrid1.ScrollIntoView(this.dataGrid1.Items[rowNo]);

      this.dataGrid1.SelectedCells.Add(cellInfo);

      m_StringResourceBuilder.SelectedGridRowIndex = rowNo;

      SetEnabled();

      m_IsSelectedCellChangedProgrammatically = false;
    }

    private StringResource GetSelectedItem()
    {
      //Debug.Print(">>> {0}", (this.dataGrid1.SelectedCells.Count == 0) ? "null" : (this.dataGrid1.SelectedCells[0].Item as StringResource).Name);

      if (this.dataGrid1.SelectedCells.Count > 0)
        return (this.dataGrid1.SelectedCells[0].Item as StringResource);
      else
        return (null);
    }

    private int GetSelectedRowIndex()
    {
      object currentItem = GetSelectedItem();

      if (currentItem == null)
        return (0);

      return (this.dataGrid1.Items.IndexOf(currentItem));
    }

    private int GetSelectedColIndex()
    {
      if (this.dataGrid1.CurrentColumn == null)
        return (m_LastCurrentColumn);

      return (this.dataGrid1.CurrentColumn.DisplayIndex);
    }

    private void ResetCurrentCellToSelection()
    {
      bool isCellChangeProgrammatically = m_IsSelectedCellChangedProgrammatically;

      m_IsSelectedCellChangedProgrammatically = true;

      if ((this.dataGrid1.SelectedCells.Count > 0) && (this.dataGrid1.CurrentCell != this.dataGrid1.SelectedCells[0]))
        this.dataGrid1.CurrentCell = this.dataGrid1.SelectedCells[0];

      m_IsSelectedCellChangedProgrammatically = isCellChangeProgrammatically;
    }

    private void GetGridSelectionStates(out bool isEmpty,
                                        out bool isFirst,
                                        out bool isLast)
    {
      int count         = this.dataGrid1.Items.Count,
          selectedIndex = GetSelectedRowIndex();

      isEmpty = (count == 0);
      isFirst = (selectedIndex == 0);
      isLast  = (selectedIndex == count - 1);
    }
    #endregion

    #endregion //Private members -------------------------------------------------------------------

    #region Public members

    //Rescan

    internal bool GetRescanEnabled()
    {
      return (true);
    }

    internal void DoRescan()
    {
      m_StringResourceBuilder.DoBrowse();
    }

    //First

    internal bool GetFirstEnabled()
    {
      bool isEmpty,
           isFirst,
           isLast;

      GetGridSelectionStates(out isEmpty, out isFirst, out isLast);

      return (!isFirst && !isEmpty);
    }

    internal void DoGotoFirst()
    {
      SelectCell(0, GetSelectedColIndex());
      m_StringResourceBuilder.SelectStringInTextDocument();
    }

    //Previous

    internal bool GetPreviousEnabled()
    {
      bool isEmpty,
           isFirst,
           isLast;

      GetGridSelectionStates(out isEmpty, out isFirst, out isLast);

      return (!isFirst && !isEmpty);
    }

    internal void DoGotoPrevious()
    {
      SelectCell(GetSelectedRowIndex() - 1, GetSelectedColIndex());
      m_StringResourceBuilder.SelectStringInTextDocument();
    }

    //Next

    internal bool GetNextEnabled()
    {
      bool isEmpty,
           isFirst,
           isLast;

      GetGridSelectionStates(out isEmpty, out isFirst, out isLast);

      return (!isLast && !isEmpty);
    }

    internal void DoGotoNext()
    {
      SelectCell(GetSelectedRowIndex() + 1, GetSelectedColIndex());
      m_StringResourceBuilder.SelectStringInTextDocument();
    }

    //Last

    internal bool GetLastEnabled()
    {
      bool isEmpty,
           isFirst,
           isLast;

      GetGridSelectionStates(out isEmpty, out isFirst, out isLast);

      return (!isLast && !isEmpty);
    }

    internal void DoGotoLast()
    {
      SelectCell(this.dataGrid1.Items.Count - 1, GetSelectedColIndex());
      m_StringResourceBuilder.SelectStringInTextDocument();
    }

    //Make

    internal bool GetMakeEnabled()
    {
      bool isEmpty,
           isFirst,
           isLast;

      GetGridSelectionStates(out isEmpty, out isFirst, out isLast);

      return (!isEmpty);
    }

    internal void DoMake()
    {
      try
      {
        m_IsMakeInProgress = true;
        m_StringResourceBuilder.BuildAndUseResource();
      }
      finally
      {
        m_IsMakeInProgress = false;
      }
    }

    //Settings

    internal bool GetSettingsEnabled()
    {
      return (true);
    }

    internal void DoSettings()
    {
      GUI.SettingsWindow window = new GUI.SettingsWindow(m_StringResourceBuilder.Settings);

      if (window.ShowDialog() ?? false)
      {
        m_StringResourceBuilder.SaveSettings();
        DoRescan();
      } //if

#if xDEBUG
      OutlineCode();
#endif
    }

    #endregion //Public members -----------------------------------------------------------------------
  } //class
} //namespace
