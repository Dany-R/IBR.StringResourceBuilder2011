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
using System.Windows.Shapes;

namespace IBR.StringResourceBuilder2011.GUI
{
  /// <summary>
  /// Interaction logic for SettingsWindow.xaml
  /// </summary>
  public partial class SettingsWindow : Window
  {
    #region Constructor

    public SettingsWindow()
    {
      InitializeComponent();
    }

    public SettingsWindow(Settings settings)
      : this()
    {
      m_Settings = settings;

      this.cbIgnoreUpToNCharactersStrings.IsChecked = m_Settings.IsIgnoreStringLength;
      this.nudIgnoreStringLength.Value              = (decimal)m_Settings.IgnoreStringLength;
      this.cbIgnoreWhiteSpaceStrings.IsChecked      = m_Settings.IsIgnoreWhiteSpaceStrings;
      this.cbIgnoreNumberStrings.IsChecked          = m_Settings.IsIgnoreNumberStrings;
      this.cbIgnoreVerbatimStrings.IsChecked        = m_Settings.IsIgnoreVerbatimStrings;
      this.cbUseGlobalResourceFile.IsChecked        = m_Settings.IsUseGlobalResourceFile;
      this.txtGlobalResourceFileName.Text           = m_Settings.GlobalResourceFileName;
      this.cbDontUseResourceAlias.IsChecked         = m_Settings.IsDontUseResourceAlias;

      this.lstIgnoreStrings.Items    = m_Settings.IgnoreStrings;
      this.lstIgnoreSubStrings.Items = m_Settings.IgnoreSubStrings;

      this.lstIgnoreMethods.Items   = m_Settings.IgnoreMethods;
      this.lstIgnoreArguments.Items = m_Settings.IgnoreMethodsArguments;

      this.lstIgnoreStrings.IsEnabled    = !m_Settings.IgnoreStrings.Contains("@@@disabled@@@");
      this.lstIgnoreSubStrings.IsEnabled = !m_Settings.IgnoreSubStrings.Contains("@@@disabled@@@");
      this.lstIgnoreMethods.IsEnabled    = !m_Settings.IgnoreMethods.Contains("@@@disabled@@@");
      this.lstIgnoreArguments.IsEnabled  = !m_Settings.IgnoreMethodsArguments.Contains("@@@disabled@@@");
    }

    #endregion //Constructor -----------------------------------------------------------------------

    #region Types
    #endregion //Types -----------------------------------------------------------------------------

    #region Fields
    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties

    private Settings m_Settings;
    public Settings Settings
    {
      get { return (m_Settings); }
      //set { m_Settings = value; }
    }

    #endregion //Properties ------------------------------------------------------------------------

    #region Events

    private void btnOK_Click(object sender, RoutedEventArgs e)
    {
      if (this.cbUseGlobalResourceFile.IsChecked ?? false)
      {
        //[12-10-03 DR]: empty for standard global resource file
        //if (string.IsNullOrEmpty(this.txtGlobalResourceFileName.Text))
        //{
        //  MessageBox.Show("The global resource file name must not be empty.",
        //                  "Settings", MessageBoxButton.OK, MessageBoxImage.Error);

        //  if (!this.tabiOptions.IsSelected)
        //    this.tabiOptions.IsSelected = true;

        //  this.txtGlobalResourceFileName.Focus();
        //  return;
        //} //if

        if (this.txtGlobalResourceFileName.Text.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0)
        {
          MessageBox.Show("The global resource file name contains at least one illegal charter.",
                          "Settings", MessageBoxButton.OK, MessageBoxImage.Error);

          if (!this.tabiOptions.IsSelected)
            this.tabiOptions.IsSelected = true;

          this.txtGlobalResourceFileName.Focus();
          return;
        } //if 
      } //if

      if (m_Settings == null)
        m_Settings = new Settings();

      m_Settings.IsIgnoreStringLength      = this.cbIgnoreUpToNCharactersStrings.IsChecked ?? false;
      m_Settings.IgnoreStringLength        = (int)this.nudIgnoreStringLength.Value;
      m_Settings.IsIgnoreWhiteSpaceStrings = this.cbIgnoreWhiteSpaceStrings.IsChecked ?? false;
      m_Settings.IsIgnoreNumberStrings     = this.cbIgnoreNumberStrings.IsChecked ?? false;
      m_Settings.IsIgnoreVerbatimStrings   = this.cbIgnoreVerbatimStrings.IsChecked ?? false;
      m_Settings.IsUseGlobalResourceFile   = this.cbUseGlobalResourceFile.IsChecked ?? false;
      m_Settings.GlobalResourceFileName    = (this.txtGlobalResourceFileName.Text ?? string.Empty).Trim();
      m_Settings.IsDontUseResourceAlias    = this.cbDontUseResourceAlias.IsChecked ?? false;

      m_Settings.IgnoreStrings.Clear();
      m_Settings.IgnoreStrings.AddRange(this.lstIgnoreStrings.Items);
      m_Settings.IgnoreSubStrings.Clear();
      m_Settings.IgnoreSubStrings.AddRange(this.lstIgnoreSubStrings.Items);

      m_Settings.IgnoreMethods.Clear();
      m_Settings.IgnoreMethods.AddRange(this.lstIgnoreMethods.Items);
      m_Settings.IgnoreMethodsArguments.Clear();
      m_Settings.IgnoreMethodsArguments.AddRange(this.lstIgnoreArguments.Items);

      this.DialogResult = true;
      this.Close();
    }

    private void btnCancel_Click(object sender, RoutedEventArgs e)
    {
      this.DialogResult = false;
      this.Close();
    }

    #endregion //Events ----------------------------------------------------------------------------

    #region Private methods
    #endregion //Private methods -------------------------------------------------------------------

    #region Public methods
    #endregion //Public methods --------------------------------------------------------------------
  } //class
}
