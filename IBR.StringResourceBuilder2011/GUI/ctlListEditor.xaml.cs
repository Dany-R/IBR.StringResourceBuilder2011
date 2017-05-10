using System;
using System.Collections.Generic;
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
using System.ComponentModel;
using System.Windows.Markup;
using System.Collections.Specialized;

namespace IBR.StringResourceBuilder2011.GUI
{
  /// <summary>
  /// Interaction logic for ctlListEditor.xaml
  /// </summary>
  [DefaultProperty("Header")]
  public partial class ctlListEditor : UserControl
  {
    #region Constructor

    public ctlListEditor()
    {
      InitializeComponent();

      this.btnUndo.IsEnabled   = false;
      this.btnRedo.IsEnabled   = false;
      this.btnAdd.IsEnabled    = false;
      this.btnRemove.IsEnabled = false;
    }

    static ctlListEditor()
    {
      HeaderProperty = DependencyProperty.Register("Header", typeof(string), typeof(ctlListEditor),
                                                   new FrameworkPropertyMetadata(">>Header<<",
                                                                                 new PropertyChangedCallback(OnHeaderChanged)));
    }

    #endregion //Constructor -----------------------------------------------------------------------

    #region Types

    private enum eAction
    {
      Add,
      Delete
    }

    private struct tAction
    {
      public tAction(eAction action,
                     string text)
      {
        Action = action;
        Text   = text;
      }

      public eAction Action;
      public string Text;
    }

    #endregion //Types -----------------------------------------------------------------------------

    #region Fields

    public static DependencyProperty HeaderProperty;

    private List<tAction> m_UndoBuffer = new List<tAction>();
    private List<tAction> m_RedoBuffer = new List<tAction>();

    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties

    public string Header
    {
      get { return ((string)GetValue(HeaderProperty)); }
      set { SetValue(HeaderProperty, value); }
    }

    public IEnumerable<string> Items
    {
      get { return (this.lstList.Items.OfType<string>()); }
      set
      {
        this.lstList.Items.Clear();

        foreach (string item in value)
          this.lstList.Items.Add(item);
      }
    }

    #endregion //Properties ------------------------------------------------------------------------

    #region Events

    private void lstList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
      HandleSelectionChanged();
    }
    
    private void txtItem_KeyDown(object sender, KeyEventArgs e)
    {
      e.Handled = HandleKeyDown(e.Key);
    }

    private void txtItem_TextChanged(object sender, TextChangedEventArgs e)
    {
      HandleTextChanged();
    }

    private void btnUndo_Click(object sender, RoutedEventArgs e)
    {
      HandleUndoRedo(true);
    }

    private void btnRedo_Click(object sender, RoutedEventArgs e)
    {
      HandleUndoRedo(false);
    }

    private void btnAdd_Click(object sender, RoutedEventArgs e)
    {
      HandleAdd();
    }

    private void btnRemove_Click(object sender, RoutedEventArgs e)
    {
      HandleRemove();
    }

    #endregion //Events ----------------------------------------------------------------------------

    #region Private methods

    private static void OnHeaderChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    {
      ctlListEditor cle = sender as ctlListEditor;

      if ((e.Property == HeaderProperty) && ((string)e.NewValue != (string)e.OldValue))
        cle.Header = (string)e.NewValue;
    }

    private void HandleSelectionChanged()
    {
      HandleTextChanged();
    }

    private bool HandleKeyDown(Key key)
    {
      bool isHandled = false;

      switch (key)
      {
        case Key.Escape:
          object selectedItem = this.lstList.SelectedItem;
          this.txtItem.Text = (selectedItem != null) ? selectedItem.ToString() : string.Empty;
          this.txtItem.SelectionStart = this.txtItem.Text.Length;

          isHandled = true;
          break;

        case Key.Return:
          if (this.btnAdd.IsEnabled)
            this.btnAdd.RaiseEvent(new RoutedEventArgs(Button.ClickEvent, this.btnAdd));
          else if (this.btnRemove.IsEnabled)
            this.btnRemove.RaiseEvent(new RoutedEventArgs(Button.ClickEvent, this.btnRemove));

          isHandled = true;
          break;

        default:
          HandleTextChanged();
          break;
      } //switch

      return (isHandled);
    }

    private void HandleTextChanged()
    {
      string text       = this.txtItem.Text;
      bool   hasText    = !string.IsNullOrEmpty(text),
             isExisting = hasText && this.lstList.Items.Contains(text);

      this.btnAdd.IsEnabled    = !isExisting && hasText;
      this.btnRemove.IsEnabled =  isExisting;
    }

    private void HandleUndoRedo(bool isUndo)
    {
      List<tAction> sourceBuffer      = isUndo ? m_UndoBuffer : m_RedoBuffer,
                    destinationBuffer = isUndo ? m_RedoBuffer : m_UndoBuffer;

      if (sourceBuffer.Count > 0)
      {
        int index      = sourceBuffer.Count - 1;
        tAction action = sourceBuffer[index];

        if ((isUndo && (action.Action == eAction.Add)) || (!isUndo && (action.Action == eAction.Delete)))
        {
          #region remove from list
          int oldIndex = this.lstList.Items.IndexOf(action.Text);

          this.lstList.Items.Remove(action.Text);

          if (this.lstList.Items.Count == 0)
            HandleTextChanged();
          else
            this.lstList.SelectedIndex = (this.lstList.Items.Count > oldIndex) ? oldIndex : this.lstList.Items.Count - 1;
          #endregion
        }
        else
        {
          #region add to list
          this.lstList.Items.Add(action.Text);
          this.lstList.SelectedIndex = this.lstList.Items.IndexOf(action.Text);
          #endregion
        } //else

        destinationBuffer.Add(action);
        sourceBuffer.RemoveAt(index);
      } //if

      this.btnUndo.IsEnabled = (m_UndoBuffer.Count > 0);
      this.btnRedo.IsEnabled = (m_RedoBuffer.Count > 0);
    }

    private void HandleAdd()
    {
      string text = this.txtItem.Text;

      this.lstList.Items.Add(text);
      this.lstList.SelectedIndex = this.lstList.Items.IndexOf(text);

      m_RedoBuffer.Clear();
      m_UndoBuffer.Add(new tAction(eAction.Add, text));

      this.btnUndo.IsEnabled = (m_UndoBuffer.Count > 0);
      this.btnRedo.IsEnabled = (m_RedoBuffer.Count > 0);
    }

    private void HandleRemove()
    {
      int oldIndex = this.lstList.SelectedIndex;

      string text = this.txtItem.Text;
      this.lstList.Items.Remove(text);

      if (this.lstList.Items.Count == 0)
        HandleTextChanged();
      else
        this.lstList.SelectedIndex = (this.lstList.Items.Count > oldIndex) ? oldIndex : this.lstList.Items.Count - 1;

      m_RedoBuffer.Clear();
      m_UndoBuffer.Add(new tAction(eAction.Delete, text));

      this.btnUndo.IsEnabled = (m_UndoBuffer.Count > 0);
      this.btnRedo.IsEnabled = (m_RedoBuffer.Count > 0);
    }

    #endregion //Private methods -------------------------------------------------------------------

    #region Public methods
    #endregion //Public methods --------------------------------------------------------------------
  } //class
} //namespace
