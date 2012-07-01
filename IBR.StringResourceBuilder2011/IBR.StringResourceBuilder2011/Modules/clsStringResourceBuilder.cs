using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

using EnvDTE;
using EnvDTE80;

using Microsoft.VisualStudio.Shell;
using System.Windows.Threading;
using System.Windows;
using System.Xml;
using System.Security.AccessControl;



namespace IBR.StringResourceBuilder2011.Modules
{
  class StringResourceBuilder
  {
    #region Constructor

    public StringResourceBuilder(DTE2 dte2)
    {
      Trace.WriteLine("Entering SRBControl()");

      m_Dte2 = dte2;

      LoadIgnoreStrings();
    }

    #endregion //Constructor -----------------------------------------------------------------------

    #region Types
    #endregion //Types -----------------------------------------------------------------------------

    #region Fields

    private DTE2 m_Dte2;

    private List<StringResource> m_StringResources = new List<StringResource>();
    private StringResource m_SelectedStringResource;

    private string m_SettingsFile;
    private bool m_SettingsPathReadOnly;
    private bool m_SettingsFileReadOnly;

    private Settings m_Settings;

    private TextDocument m_TextDocument;
    private bool m_IsCSharp;

    private bool m_IsTextMoveSuspended;
    //private bool m_IsLineChanged;

    private bool m_IsMakePerformed;

    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties

    public Action<double> InitProgress { get; set; }
    public Action HideProgress { get; set; }
    public Action<int> SetProgress { get; set; }

    public Action ClearGrid { get; set; }
    public Action<System.Collections.IEnumerable> SetGridItemsSource { get; set; }
    public Action RefreshGrid { get; set; }
    public Action<int, int> SelectCell { get; set; }
    public Func<StringResource> GetSelectedItem { get; set; }
    public Func<int> GetSelectedColIndex { get; set; }
    public Func<int> GetSelectedRowIndex { get; set; }

    private EnvDTE.Window m_Window;
    public EnvDTE.Window Window
    {
      get { return (m_Window); }
      set { m_Window = value; }
    }

    private EnvDTE.Window m_FocusedTextDocumentWindow;
    public EnvDTE.Window FocusedTextDocumentWindow
    {
      get { return (m_FocusedTextDocumentWindow); }
      set { m_FocusedTextDocumentWindow = value; }
    }

    private int m_OldCurrentLine;
    public int OldCurrentLine
    {
      get { return (m_OldCurrentLine); }
      set { m_OldCurrentLine = value; }
    }

    private bool m_IsBrowsing;
    public bool IsBrowsing
    {
      get { return (m_IsBrowsing); }
      set { m_IsBrowsing = value; }
    }

    private int m_SelectedRowIndex;
    public int SelectedRowIndex
    {
      get { return (m_SelectedRowIndex); }
      set { m_SelectedRowIndex = value; }
    }

    public Settings Settings
    {
      get { return (m_Settings); }
      private set { m_Settings = value; }
    }

    #endregion //Properties ------------------------------------------------------------------------

    #region Events

    #endregion //Events ----------------------------------------------------------------------------

    #region Private methods

    private void SuspendTextMove()
    {
      m_IsTextMoveSuspended = true;
    }

    private void ResumeTextMove()
    {
      m_IsTextMoveSuspended = false;
    }

#if DEBUG
    System.Diagnostics.Stopwatch m_sw;
#endif

    private void BrowseForStrings(EnvDTE.Window window)
    {
      if (m_IsBrowsing)
        return;

      Debug.Print("\n> BrowseForStrings()");

      m_IsBrowsing = true;

      this.ClearGrid();

      m_StringResources.Clear();

      if (!IsSourceCode(window))
      {
        Debug.Print("> BrowseForStrings() - cancelled");
        m_IsBrowsing = false;
        return;
      } //if

      m_OldCurrentLine = m_TextDocument.Selection.ActivePoint.Line;

      //OutlineCode();

      this.InitProgress((double)m_TextDocument.EndPoint.Line);

#if DEBUG
      m_sw = System.Diagnostics.Stopwatch.StartNew();
#endif

#if notyet
      if (!m_Window.Caption.EndsWith(".xaml"))
        Parser.ParseForStrings(m_Window, m_StringResources, m_Settings, this.SetProgress, ParsingCompleted);
      else
        XamlParser.ParseForStrings(m_Window, m_StringResources, m_Settings, this.SetProgress, ParsingCompleted);
#else
      Parser.ParseForStrings(m_Window, m_StringResources, m_Settings, this.SetProgress, ParsingCompleted);
#endif

      Debug.Print("> BrowseForStrings() - ended\n");
    }

    private void ParsingCompleted()
    {
#if DEBUG
      m_sw.Stop(); System.Diagnostics.Debug.Print("{0} elapsed in {1} ms", m_StringResources.Count, m_sw.Elapsed.TotalMilliseconds);
#endif

      this.HideProgress();

      this.SetGridItemsSource(m_StringResources);

      if (m_StringResources.Count > 0)
      {
        if (m_IsTextMoveSuspended)
        {
          if (m_SelectedRowIndex >= m_StringResources.Count)
            m_SelectedRowIndex = m_StringResources.Count - 1;

          this.SelectCell(m_SelectedRowIndex, this.GetSelectedColIndex());
        }
        else if (m_IsMakePerformed)
        {
          m_IsMakePerformed = false;
          SelectStringInTextDocument();
        }
        else
        {
          SelectNearestGridRow();
        } //else
      } //if

      if ((m_Window != null) && (m_Window != m_Dte2.ActiveWindow))
        m_Window.Activate();

      m_IsBrowsing = false;
    }

    private List<StringResource> GetAllStringResourcesInLine(int lineNo)
    {
      List<StringResource> stringResources = m_StringResources.FindAll(sr => (sr.Location.X == lineNo));
      return (stringResources);
    }

    private StringResource GetLastStringResourceBeforeLine(int lineNo)
    {
      StringResource stringResource = m_StringResources.FindLast(sr => (sr.Location.X < lineNo));
      return (stringResource);
    }

    private StringResource GetFirstStringResourceAfterLine(int lineNo)
    {
      StringResource stringResource = m_StringResources.Find(sr => (sr.Location.X > lineNo));
      return (stringResource);
    }

    #region Source code
    private static bool IsLanguageSupported(EnvDTE.Document document)
    {
      if (document == null)
        return (false);

      if ((document.Language != "CSharp") && (document.Language != "Basic"))
        return (false);

      return (true);
    }

    private bool IsSourceCode(EnvDTE.Window window)
    {
      Debug.Print(">>>> IsSourceCode() - window='{0}' active='{1}'", (window == null) ? "null" : window.Caption, (m_Dte2.ActiveDocument == null) ? "null" : m_Dte2.ActiveDocument.ActiveWindow.Caption);

      m_TextDocument = null;

      m_Window = window;

      if (m_Window == null)
      {
        if (!IsLanguageSupported(m_Dte2.ActiveDocument))
          return (false);

        //foreach (EnvDTE.Window w in m_Dte2.ActiveDocument.Windows)
        //{
        //  if (w.Caption.EndsWith("[Design]", StringComparison.OrdinalIgnoreCase))
        //    continue;

        //  m_Window = w;
        //  break;
        //} //foreach
        if ((m_Dte2.ActiveDocument.ActiveWindow != null)
            && !m_Dte2.ActiveDocument.ActiveWindow.Caption.EndsWith("[Design]", StringComparison.OrdinalIgnoreCase))
          m_Window = m_Dte2.ActiveDocument.ActiveWindow;

        if (m_Window == null)
          return (false);
      }
      else
      {
        if (!IsLanguageSupported(m_Window.Document))
          return (false);

        if (m_Window.Caption.EndsWith("[Design]", StringComparison.OrdinalIgnoreCase))
          return (false);
      } //else

      m_IsCSharp     = (m_Window.Document.Language == "CSharp");
      m_TextDocument = m_Window.Document.Object("TextDocument") as TextDocument;
      if (m_TextDocument == null)
        return (false);

      System.Diagnostics.Debug.Print("m_TextDocument.Selection.ActivePoint.Line = {0}",
                                     m_TextDocument.Selection.ActivePoint.Line);

      return (true);
    }

    private bool Find(EditPoint2 editPoint,
                      string text)
    {
      EditPoint endPoint = null;
      TextRanges tags = null;
      return (editPoint.FindPattern(text, (int)vsFindOptions.vsFindOptionsMatchInHiddenText,
                                    ref endPoint, ref tags));
    }
    #endregion

    #region Resource
    private ProjectItem OpenOrCreateResourceFile()
    {
      ProjectItem prjItem = null;

      string resourceFileName = System.IO.Path.ChangeExtension(m_Window.ProjectItem.Name, "Resources.resx");

      //get the projects project-items collection
      ProjectItems prjItems = m_Window.Project.ProjectItems;

      try
      {
        //try to get the parent project-items collection (if in a sub folder)
        prjItems = ((ProjectItem)m_Window.ProjectItem.Collection.Parent).ProjectItems;
      }
      catch { }

      try
      {
        prjItem = prjItems.Item(resourceFileName);
      }
      catch { }

      if (prjItem == null)
      {
        #region not in project but file exists? -> ask user if delete
        string projectPath      = System.IO.Path.GetDirectoryName(m_Window.ProjectItem.ContainingProject.FullName),
               //resourceFilePath = System.IO.Path.GetDirectoryName(m_Window.ProjectItem.FileNames[0]),
               resourceFile     = System.IO.Path.ChangeExtension(m_Window.ProjectItem.FileNames[0], "Resources.resx"),
               designerFile     = System.IO.Path.ChangeExtension(resourceFile, ".Designer." + ((m_IsCSharp) ? "cs" : "vb"));

        if (System.IO.File.Exists(resourceFile) || System.IO.File.Exists(designerFile))
        {
          string msg = string.Format("The resource file already exists though it is not included in the project:\r\n\r\n"
                                     + "'{0}'\r\n\r\n"
                                     + "Do you want to overwrite the existing resource file?",
                                     resourceFile.Substring(projectPath.Length).TrimStart('\\'));

          if (MessageBox.Show(msg, "Make resource", MessageBoxButton.YesNo,
                              MessageBoxImage.Question) != MessageBoxResult.Yes)
            return (null);
          else
          {
            TryToSilentlyDeleteIfExistsEvenIfReadOnly(resourceFile);
            TryToSilentlyDeleteIfExistsEvenIfReadOnly(designerFile);
          } //else
        } //if
        #endregion

        try
        {
          string language = (m_IsCSharp) ? "csproj" : "vbproj";

          // Retrieve the path to the resource template.
          string itemPath = ((Solution2)m_Dte2.Solution).GetProjectItemTemplate("Resource.zip", language);

          //create a new project item based on the template
          prjItem = prjItems.AddFromTemplate(itemPath, resourceFileName); //returns always null ...
          prjItem = prjItems.Item(resourceFileName);
        }
        catch (Exception ex)
        {
          Trace.WriteLine(string.Format("### OpenOrCreateResourceFile() - {0}", ex.ToString()));
          return (null);
        }
      } //if

      //open the resx file
      if (!prjItem.IsOpen[Constants.vsViewKindAny])
        prjItem.Open(Constants.vsViewKindDesigner);

      return (prjItem);
    }

    private void BuildAndUseResource(ProjectItem prjItem)
    {
      try
      {
        string resxFile = prjItem.Document.FullName;

        string nameSpace = m_Window.Project.Properties.Item("RootNamespace").Value.ToString();

        #region get namespace from ResX designer file
        foreach (ProjectItem item in prjItem.ProjectItems)
        {
          if (item.Name.EndsWith(".resx", StringComparison.InvariantCultureIgnoreCase))
            continue;

          foreach (CodeElement element in item.FileCodeModel.CodeElements)
          {
            if (element.Kind == vsCMElement.vsCMElementNamespace)
            {
              nameSpace = element.FullName;
              break;
            } //if
          } //foreach
        } //foreach
        #endregion

        //close the resx file to modify it
        prjItem.Document.Close(vsSaveChanges.vsSaveChangesYes);

        string name,
               value,
               comment = string.Empty;
        StringResource stringResource = m_SelectedStringResource;// this.dataGrid1.CurrentItem as StringResource;

        name = stringResource.Name;
        value = stringResource.Text;

        if (!m_IsCSharp || (m_TextDocument.Selection.Text.Length > 0) && (m_TextDocument.Selection.Text[0] == '@'))
          value = value.Replace("\"\"", "\\\"");

        if (!AppendStringResource(resxFile, name, value, comment))
          return;

        //create the designer class
        VSLangProj.VSProjectItem vsPrjItem = prjItem.Object as VSLangProj.VSProjectItem;
        if (vsPrjItem != null)
          vsPrjItem.RunCustomTool();

        int replaceLength = m_TextDocument.Selection.Text.Length;
        if (replaceLength > 0)
        {
          string className = System.IO.Path.GetFileNameWithoutExtension(resxFile).Replace('.', '_');
          string aliasName = className.Substring(0, className.Length - 6); //..._Resources -> ..._Res
          //string resourceCall = string.Format("global::{0}.{1}.{2}", nameSpace, className, name);
          string resourceCall = string.Format("{0}.{1}", aliasName, name);

          int oldRow = m_TextDocument.Selection.ActivePoint.Line;

          CheckAndAddAlias(nameSpace, className, aliasName);
          m_TextDocument.Selection.Text = resourceCall;

          UpdateTableAndSelectNext(resourceCall.Length - replaceLength, m_TextDocument.Selection.ActivePoint.Line - oldRow);
        } //if
      }
      catch (Exception ex)
      {
        Trace.WriteLine(string.Format("### BuildAndUseResource() - {0}", ex.ToString()));
      }
    }

    private void CheckAndAddAlias(string nameSpace,
                                  string className,
                                  string aliasName)
    {
      string resourceAlias1 = string.Format("using {0} = ", aliasName),
             resourceAlias2 = string.Format("global::{0}.{1}", nameSpace, className);

      if (!m_IsCSharp)
      {
        resourceAlias1 = string.Format("Imports {0} = ", aliasName);
        resourceAlias2 = string.Format("{0}.{1}", nameSpace, className);
      } //if

      CodeElements elements   = m_TextDocument.Parent.ProjectItem.FileCodeModel.CodeElements;
      CodeElement lastElement = null;
      bool isImport           = false;

      #region find alias or last using/Import element (return if alias found)
      foreach (CodeElement element in elements) //not really fast but more safe
      {
        if (!isImport)
        {
          //find first using/import statement
          if (element.Kind == vsCMElement.vsCMElementImportStmt)
            isImport = true;
        }
        else
        {
          //using/import statement was available so find next NON using/import statement
          if (element.Kind != vsCMElement.vsCMElementImportStmt)
            break;
        } //else

        if (element.Kind == vsCMElement.vsCMElementOptionStmt)
          //save last option statement
          lastElement = element;
        else if (element.Kind == vsCMElement.vsCMElementImportStmt)
        {
          //save last using/import statement
          lastElement = element;

          //check if resource alias is already there
          CodeImport importElement = element as CodeImport;
          if ((importElement.Alias != null) && importElement.Alias.Equals(aliasName) && importElement.Namespace.Equals(resourceAlias2))
            return;
        } //if
      } //foreach
      #endregion

      EditPoint insertPoint = null;

      if (lastElement == null)
        insertPoint = m_TextDocument.CreateEditPoint(m_TextDocument.StartPoint); //beginning of text
      else
      {
        //behind last element
        insertPoint = lastElement.EndPoint.CreateEditPoint();
        insertPoint.LineDown(1);
        insertPoint.StartOfLine();

        if (lastElement.Kind == vsCMElement.vsCMElementOptionStmt)
          insertPoint.Insert(Environment.NewLine);
      } //else

      if (m_IsCSharp)
        resourceAlias2 += ";";

      string alias = resourceAlias1 + resourceAlias2 + Environment.NewLine;

      insertPoint.Insert(alias);
    }

    private bool AppendStringResource(string resxFileName,
                                      string name,
                                      string value,
                                      string comment)
    {
      try
      {
        XmlElement dataElement,
                   valueElement;
        XmlAttribute attribute;

        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(resxFileName);

        string xmlPath = string.Format("descendant::data[(attribute::name='{0}')]/descendant::value",
                                       name);
        XmlNode xmlNode = xmlDoc.DocumentElement.SelectSingleNode(xmlPath);
        if (xmlNode != null)
        {
          //xmlNode = xmlNode.SelectSingleNode("descendant::value");
          string msg = string.Format("This resource name already exists:\r\n\r\n"
                                     + "{0} = '{1}'\r\n"
                                     + "(new string is '{2}')\r\n\r\n"
                                     + "Do you want to use the existing resource instead?",
                                     name, xmlNode.InnerText, value);
          if (MessageBox.Show(msg, "Make resource", MessageBoxButton.YesNo,
                              MessageBoxImage.Question) != MessageBoxResult.Yes)
            return (false);
          else
            return (true);
        } //if

        if (m_IsCSharp)
        {
          if (value.Contains("\\r"))
            value = value.Replace("\\r", "\r");

          if (value.Contains("\\n"))
            value = value.Replace("\\n", "\n");

          if (value.Contains("\\t"))
            value = value.Replace("\\t", "\t");

          if (value.Contains("\\0"))
            value = value.Replace("\\0", "\0");

          if (value.Contains(@"\\"))
            value = value.Replace(@"\\", @"\");
        } //if

        if (value.Contains("\\\""))
          value = value.Replace("\\\"", "\"");

        // <data name="eMsgBoxButtons_RetryCancel" xml:space="preserve">
        //   <value>&amp;Wiederholen;&amp;Abbruch</value>
        //   <comment>&amp;Retry;&amp;Cancel</comment>
        // </data>

        dataElement = xmlDoc.CreateElement("data");
        {
          attribute = xmlDoc.CreateAttribute("name");
          attribute.Value = name;
          dataElement.Attributes.Append(attribute);

          attribute = xmlDoc.CreateAttribute("xml:space");
          attribute.Value = "preserve";
          dataElement.Attributes.Append(attribute);

          valueElement = xmlDoc.CreateElement("value");
          valueElement.InnerText = value;
          dataElement.AppendChild(valueElement);

          if (!string.IsNullOrEmpty(comment))
          {
            valueElement = xmlDoc.CreateElement("comment");
            valueElement.InnerText = comment;
            dataElement.AppendChild(valueElement);
          } //if
        }
        xmlDoc.DocumentElement.AppendChild(dataElement);

        xmlDoc.Save(resxFileName);
      }
      catch (Exception ex)
      {
        Trace.WriteLine(string.Format("### AppendStringResource() - {0}", ex.ToString()));
        return (false);
      }

      return (true);
    }

    private void RemoveStringResource(int index)
    {
      m_StringResources.RemoveAt(index);
    }
    #endregion

    #region Ignore strings
    private void LoadIgnoreStrings()
    {
      string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
      System.Diagnostics.Debug.Print(path);

      FileSystemRights rights = CUtil.GetCurrentUsersFileSystemRights(path);
      System.Diagnostics.Debug.Print("-> {0}", rights);

      m_SettingsPathReadOnly = !CUtil.Contains(rights, FileSystemRights.Write);
      m_SettingsFile         = System.IO.Path.Combine(path, "Settings.xml");
      m_SettingsFileReadOnly = CUtil.IsFileReadOnly(m_SettingsFile) || m_SettingsPathReadOnly;

      if (System.IO.File.Exists(m_SettingsFile))
      {
        string xml = System.IO.File.ReadAllText(m_SettingsFile, UTF8Encoding.UTF8);
        m_Settings = Settings.DeSerialize(xml);
      }
      else
      {
        //defaults
        m_Settings = new Settings();
        m_Settings.IsIgnoreWhiteSpaceStrings = true;
        m_Settings.IsIgnoreNumberStrings     = true;
        m_Settings.IsIgnoreStringLength      = true;
        m_Settings.IgnoreStringLength        = 2;
      } //else

#if !IGNORE_METHOD_ARGUMENTS
      if (m_Settings != null)
        m_Settings.IgnoreMethodsArguments.Add("@@@disabled@@@");
#endif
    }

    private void SaveIgnoreStrings()
    {
      if (m_SettingsPathReadOnly)
        return;

      if (m_Settings == null)
        return;

      if (m_SettingsFileReadOnly)
        return;

#if !IGNORE_METHOD_ARGUMENTS
      m_Settings.IgnoreMethodsArguments.Remove("@@@disabled@@@");
#endif

      System.IO.File.WriteAllText(m_SettingsFile, m_Settings.Serialize(), UTF8Encoding.UTF8);

#if !IGNORE_METHOD_ARGUMENTS
      m_Settings.IgnoreMethodsArguments.Add("@@@disabled@@@");
#endif
    }
    #endregion

    #region Grid
    private void SelectNearestGridRow()
    {
      int curLine = m_TextDocument.Selection.ActivePoint.Line,
          curCol  = m_TextDocument.Selection.ActivePoint.LineCharOffset,
          rowNo   = -1,
          lineNo  = -1;

      List<StringResource> stringResources = GetAllStringResourcesInLine(curLine);

      if (stringResources.Count > 0)
      {
        if (stringResources[stringResources.Count - 1].Location.Y < curCol)
        {
          //same location line (same or previous string)
          StringResource stringResource = stringResources[stringResources.Count - 1];

          rowNo  = m_StringResources.IndexOf(stringResource);
          lineNo = stringResource.Location.X;
          Debug.Print("> Found at ({0}, {1}): '{2}'", lineNo, stringResource.Location.Y, stringResource.Text);
        }
        else
        {
          //same location line (same or next string)
          foreach (StringResource stringResource in stringResources)
          {
            int y = stringResource.Location.Y;

            if (((y <= curCol) && (stringResource.Text.Length + y + 2 >= curCol))
                || (y > curCol))
            {
              rowNo  = m_StringResources.IndexOf(stringResource);
              lineNo = stringResource.Location.X;
              Debug.Print("> Found at ({0}, {1}): '{2}'", lineNo, stringResource.Location.Y, stringResource.Text);
              break;
            } //if
          } //foreach
        } //else
      }
      else
      {
        //find nearest line (select only in grid)
        StringResource stringResource1 = GetLastStringResourceBeforeLine(curLine),
                       stringResource2 = GetFirstStringResourceAfterLine(curLine);

        if ((stringResource1 != null)
            && ((stringResource2 == null)
                || ((curLine - stringResource1.Location.X) <= (stringResource2.Location.X - curLine))))
        {
          rowNo  = m_StringResources.IndexOf(stringResource1);
          lineNo = stringResource1.Location.X;
          Debug.Print("> Found at ({0}, {1}): '{2}'", lineNo, stringResource1.Location.Y, stringResource1.Text);
        }
        else if (stringResource2 != null)
        {
          rowNo  = m_StringResources.IndexOf(stringResource2);
          lineNo = stringResource2.Location.X;
          Debug.Print("> Found at ({0}, {1}): '{2}'", lineNo, stringResource2.Location.Y, stringResource2.Text);
        } //else
      } //else

      if (rowNo >= 0)
      {
        this.SelectCell(rowNo, this.GetSelectedColIndex());

        if (lineNo == curLine)
          SelectStringInTextDocument();
      } //if
    }

    private void UpdateTableAndSelectNext(int deltaLength,
                                          int deltaLine)
    {
//       bool isThereAnotherText          = false;
//       string oldText                   = m_SelectedStringResource.Text;
      System.Drawing.Point oldLocation = m_SelectedStringResource.Location;

//       //check if there is another occurrence
//       foreach (StringResource stringResource in m_StringResources)
//       {
//         if (stringResource.Location.Equals(oldLocation))
//           continue;
// 
//         isThereAnotherText = (stringResource.Text == oldText);
//         if (isThereAnotherText)
//           break;
//       } //foreach

//       if (isThereAnotherText)
//       {
//         #region rebuild table
//         int oldRowNo = m_SelectedRowIndex,
//             oldColNo = this.GetSelectedColIndex();
// 
//         btnRescan_Click(this.btnRescan, new RoutedEventArgs());
// 
//         //if (m_StringResources.Count <= oldRowNo)
//         //  oldRowNo = m_StringResources.Count - 1;
//         if (this.dataGrid1.Items.Count <= oldRowNo)
//           oldRowNo = this.dataGrid1.Items.Count - 1;
// 
//         if (oldRowNo >= 0)
//           this.SelectCell(oldRowNo, oldColNo);
//         #endregion
//       }
//       else
      {
        #region remove entry and goto next
//         //for (int rowIndex = 0; rowIndex < m_StringResources.Count; ++rowIndex)
//         for (int rowIndex = 0; rowIndex < this.dataGrid1.Items.Count; ++rowIndex)
//         {
//           System.Drawing.Point loc = GetStringLocation(rowIndex);
// 
//           if (loc.Equals(oldLocation))
//             continue;
// 
//           if (loc.X > oldLocation.X)
//             break;
// 
//           if (loc.X < oldLocation.X)
//             continue;
// 
//           if (loc.Y < oldLocation.Y)
//             continue;
// 
//           //loc.Y += deltaLength;
//           m_StringResources[rowIndex].Offset(deltaLine, deltaLength);
//         } //for

        if ((deltaLength != 0) || (deltaLine != 0))
        {
          StringResource stringResource = null;

          if (m_SelectedRowIndex < m_StringResources.Count - 1)
            stringResource = m_StringResources[m_SelectedRowIndex + 1] as StringResource;
          else if (m_SelectedRowIndex > 0)
            stringResource = m_StringResources[m_SelectedRowIndex - 1] as StringResource;

          if ((deltaLength != 0) && (stringResource.Location.X == oldLocation.X) && (stringResource.Location.Y > oldLocation.Y))
            stringResource.Offset(0, deltaLength);

          if (deltaLine != 0)
            stringResource.Offset(deltaLine, 0);
        } //if

        this.ClearGrid();
        RemoveStringResource(m_SelectedRowIndex);
        this.RefreshGrid();
        this.SelectCell(m_SelectedRowIndex, this.GetSelectedColIndex());
        #endregion
      } //else

      SelectStringInTextDocument();
    }

    private System.Drawing.Point GetStringLocation()
    {
      return (this.GetSelectedItem().Location); //X==line number, Y==column
    }

    //private System.Drawing.Point GetStringLocation(int rowNo)
    //{
    //  return (m_StringResources[rowNo].Location); //X==line number, Y==column
    //}
    #endregion

    private static void TryToSilentlyDeleteIfExistsEvenIfReadOnly(string file)
    {
      if (System.IO.File.Exists(file))
      {
        System.IO.FileAttributes attribs = System.IO.File.GetAttributes(file);

        if ((attribs & System.IO.FileAttributes.ReadOnly) != 0)
          System.IO.File.SetAttributes(file, attribs & ~System.IO.FileAttributes.ReadOnly);

        try { System.IO.File.Delete(file); }
        catch { }
      } //if
    }

    [System.Diagnostics.Conditional("DEBUG")]
    private void OutlineCode()
    {
      FileCodeModel fileCM  = m_Dte2.ActiveDocument.ProjectItem.FileCodeModel;
      CodeElements elements = fileCM.CodeElements;

      System.Diagnostics.Debug.Print("about to walk top-level code elements ...");

      foreach (CodeElement element in elements)
        CollapseElement(element);
    }

    [System.Diagnostics.Conditional("DEBUG")]
    private void CollapseElement(CodeElement element)
    {
      if (element.IsCodeType && (element.Kind != vsCMElement.vsCMElementDelegate))
      {
        System.Diagnostics.Debug.Print("got type but not a delegate, named : {0} in line {1} to {2}", element.Name, element.StartPoint.Line, element.EndPoint.Line);
        CodeType ct = element as EnvDTE.CodeType;
        CodeElements mems = ct.Members;
        foreach (CodeElement melt in mems)
        {
          CollapseElement(melt);
        } //foreach
      }
      else if (element.Kind == vsCMElement.vsCMElementNamespace)
      {
        System.Diagnostics.Debug.Print("got a namespace, named: {0} in line {1} to {2}", element.Name, element.StartPoint.Line, element.EndPoint.Line);
        CodeNamespace cns = element as EnvDTE.CodeNamespace;
        System.Diagnostics.Debug.Print("set cns = elt, named: {0}", cns.Name);

        CodeElements mems_ns = cns.Members;
        if (mems_ns.Count > 0)
        {
          System.Diagnostics.Debug.Print("got cns.members");
          foreach (CodeElement melt in mems_ns)
          {
            CollapseElement(melt);
          } //foreach
          System.Diagnostics.Debug.Print("end of cns.members");
        } //if
      }
      else if (element.Kind == vsCMElement.vsCMElementFunction)
      {
        System.Diagnostics.Debug.Print("got a function, named: {0} in line {1} to {2}", element.Name, element.StartPoint.Line, element.EndPoint.Line);
        CodeFunction cf = element as EnvDTE.CodeFunction;
        System.Diagnostics.Debug.Print("set cf = elt, named: {0}, fkind: {1}", cf.Name, cf.FunctionKind);

        CodeElements mems_f = cf.Children;
        if (mems_f.Count > 0)
        {
          System.Diagnostics.Debug.Print("got cf.members");
          foreach (CodeElement melt in mems_f)
          {
            CollapseElement(melt);
          } //foreach
          System.Diagnostics.Debug.Print("end of cf.members");
        } //if
      }
      else if (element.Kind == vsCMElement.vsCMElementProperty)
      {
        System.Diagnostics.Debug.Print("got a property, named: {0} in line {1} to {2}", element.Name, element.StartPoint.Line, element.EndPoint.Line);
        CodeProperty cp = element as EnvDTE.CodeProperty;
        System.Diagnostics.Debug.Print("set cp = elt, named: {0}", cp.Name);

        CodeElements mems_p = cp.Children;
        if (mems_p.Count > 0)
        {
          System.Diagnostics.Debug.Print("got cp.members");
          foreach (CodeElement melt in mems_p)
          {
            CollapseElement(melt);
          } //foreach
          System.Diagnostics.Debug.Print("end of cp.members");
        } //if
      }
      else if (element.Kind == vsCMElement.vsCMElementVariable)
      {
        System.Diagnostics.Debug.Print("got a variable, named: {0} in line {1} to {2}", element.Name, element.StartPoint.Line, element.EndPoint.Line);
        CodeVariable cv = element as EnvDTE.CodeVariable;
        System.Diagnostics.Debug.Print("set cv = elt, named: {0}", cv.Name);

        CodeElements mems_v = cv.Children;
        if (mems_v.Count > 0)
        {
          System.Diagnostics.Debug.Print("got cv.members");
          foreach (CodeElement melt in mems_v)
          {
            CollapseElement(melt);
          } //foreach
          System.Diagnostics.Debug.Print("end of cv.members");
        } //if
      }
      else
        System.Diagnostics.Debug.Print("kind = {0} in line {1} to {2}", element.Kind, element.StartPoint.Line, element.EndPoint.Line);
    }

    #endregion //Private methods -------------------------------------------------------------------

    #region Public methods

    public void DoBrowse(bool isLineChanged)
    {
      if (isLineChanged)
        SuspendTextMove();

      BrowseForStrings(m_FocusedTextDocumentWindow);

      ResumeTextMove();
    }

    public void SelectStringInTextDocument()
    {
      StringResource stringResource = this.GetSelectedItem();

      if (stringResource == null)
        return;

      string text = stringResource.Text;
      System.Drawing.Point location = GetStringLocation();
      //bool isAtString = false;

      m_TextDocument.Selection.MoveToLineAndOffset(location.X, location.Y, false);

      if (location.Y > 1)
      {
        m_TextDocument.Selection.MoveToLineAndOffset(location.X, location.Y - 1, false);
        m_TextDocument.Selection.CharRight(true, 1);
        if (m_TextDocument.Selection.Text[0] != '@')
          m_TextDocument.Selection.MoveToLineAndOffset(location.X, location.Y, false);
        //else
        //  isAtString = true;
      } //if

      //from this point 'text' won't be valid anymore (changed for length)
      //if (!isAtString)
      //{
      //  if (text.Contains(@"\"))
      //    text = text.Replace(@"\", "#");
      //  //if (text.Contains(@"\\"))
      //  //  text = text.Replace(@"\\", "##");
      //  //if (text.Contains(@"\r"))
      //  //  text = text.Replace(@"\r", "##");
      //  //if (text.Contains(@"\n"))
      //  //  text = text.Replace(@"\n", "##");
      //  //if (text.Contains(@"\t"))
      //  //  text = text.Replace(@"\t", "##");
      //  //if (text.Contains(@"\0"))
      //  //  text = text.Replace(@"\0", "##");
      //  //if (text.Contains("\""))
      //  //  text = text.Replace("\"", "##");
      //}
      ////else
      ////{
      ////  if (text.Contains("\""))
      ////    text = text.Replace("\"", "##");
      ////  if (text.Contains(@"\\"))
      ////    text = text.Replace(@"\\", "##");
      ////} //else

      m_TextDocument.Selection.MoveToLineAndOffset(location.X, location.Y + text.Length + 2, true);

      m_SelectedStringResource = this.GetSelectedItem();
      m_SelectedRowIndex       = this.GetSelectedRowIndex();

      if ((m_Window != null) && (m_Window != m_Dte2.ActiveWindow))
        m_Window.Activate();
    }

    public void ClearStringResources()
    {
      m_StringResources.Clear();
    }

    public void BuildAndUseResource()
    {
      m_IsMakePerformed = true;

      SelectStringInTextDocument();

      ProjectItem prjItem = OpenOrCreateResourceFile();

      if (prjItem == null)
        return;

      BuildAndUseResource(prjItem);

      //SetEnabled();

      //MarkStringInTextDocument();
    }

    public void SaveSettings()
    {
      SaveIgnoreStrings();
    }

    #endregion //Public methods --------------------------------------------------------------------
  } //class
} //namespace
