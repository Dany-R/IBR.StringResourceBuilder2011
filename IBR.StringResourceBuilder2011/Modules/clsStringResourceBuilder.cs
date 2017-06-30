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

    private TextDocument m_TextDocument;
    private int m_LastDocumentLength;
    private bool m_IsCSharp;
    private string m_ProjectExtension;

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
      set
      {
        m_FocusedTextDocumentWindow = value;

        TextDocument txtDoc = (value == null) ? null : value.Document.Object("TextDocument") as TextDocument;
        m_LastDocumentLength = (txtDoc == null) ? 0 : txtDoc.EndPoint.Line;
      }
    }

    private bool m_IsBrowsing;
    public bool IsBrowsing
    {
      get { return (m_IsBrowsing); }
      set { m_IsBrowsing = value; }
    }

    private int m_SelectedGridRowIndex;
    public int SelectedGridRowIndex
    {
      get { return (m_SelectedGridRowIndex); }
      set { m_SelectedGridRowIndex = value; }
    }

    private Settings m_Settings;
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

    private void BrowseForStrings(EnvDTE.Window window,
                                  TextPoint startPoint,
                                  TextPoint endPoint)
    {
      Debug.Print("\n> BrowseForStrings()");

      if (m_IsBrowsing)
      {
        Debug.Print("> BrowseForStrings() - cancelled (already running)");
        return;
      } //if

      m_IsBrowsing = true;

      bool isFullDocument = startPoint.AtStartOfDocument && endPoint.AtEndOfDocument;

      if (isFullDocument)
        ClearStringResources();

      //OutlineCode();

      this.InitProgress((double)(endPoint.Line - startPoint.Line + 1));

#if DEBUG
      m_sw = System.Diagnostics.Stopwatch.StartNew();
#endif

#if notyet
      if (!m_Window.Caption.EndsWith(".xaml"))
        Parser.ParseForStrings(m_Window, m_StringResources, m_Settings, this.SetProgress, ParsingCompleted,
                               startPoint, endPoint, m_LastDocumentLength);
      else
        XamlParser.ParseForStrings(m_Window, m_StringResources, m_Settings, this.SetProgress, ParsingCompleted,
                                   startPoint, endPoint, m_LastDocumentLength);
#else
      Parser.ParseForStrings(m_Window, m_StringResources, m_Settings, this.SetProgress, ParsingCompleted,
                             startPoint, endPoint, m_LastDocumentLength);
#endif

      //store for next operation (moving locations)
      m_LastDocumentLength = m_TextDocument.EndPoint.Line;

      Debug.Print("> BrowseForStrings() - ended\n");
    }

    private void ParsingCompleted(bool isChanged)
    {
#if DEBUG
      m_sw.Stop(); System.Diagnostics.Debug.Print("{0} elapsed in {1} ms", m_StringResources.Count, m_sw.Elapsed.TotalMilliseconds);
#endif

      this.HideProgress();

      this.SetGridItemsSource(m_StringResources);
      this.RefreshGrid();

      if (isChanged && (m_StringResources.Count > 0))
      {
        if (m_IsTextMoveSuspended)
        {
          if (m_SelectedGridRowIndex >= m_StringResources.Count)
            m_SelectedGridRowIndex = m_StringResources.Count - 1;

          this.SelectCell(m_SelectedGridRowIndex, this.GetSelectedColIndex());
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

    private StringResource GetFirstStringResourceBehindLine(int lineNo)
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

      m_IsCSharp         = (m_Window.Document.Language == "CSharp");
      m_ProjectExtension = System.IO.Path.GetExtension(m_Window.Project.FullName).Substring(1);
      m_TextDocument     = m_Window.Document.Object("TextDocument") as TextDocument;
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
      ProjectItem resourceFilePrjItem = null;

      string resourceFileName = null,
             resourceFileDir  = null;

      if (!m_Settings.IsUseGlobalResourceFile)
      {
        resourceFileName = System.IO.Path.ChangeExtension(m_Window.ProjectItem.Name, "Resources.resx"); //file name only
        resourceFileDir  = System.IO.Path.GetDirectoryName(m_Window.ProjectItem.FileNames[0]);
      }
      else
      {
        resourceFileName = m_ProjectExtension.StartsWith("cs", StringComparison.OrdinalIgnoreCase) ? "Properties" : "My Project";
        resourceFileDir  = System.IO.Path.GetDirectoryName(m_Window.ProjectItem.ContainingProject.FullName);
      } //else

      //get the projects project-items collection
      ProjectItems prjItems = m_Window.Project.ProjectItems;

      if (!m_Settings.IsUseGlobalResourceFile)
      {
        try
        {
          //try to get the parent project-items collection (if in a sub folder)
          prjItems = ((ProjectItem)m_Window.ProjectItem.Collection.Parent).ProjectItems;
        }
        catch { }
      } //if

      try
      {
        resourceFilePrjItem = prjItems?.Item(resourceFileName);
      }
      catch { }

      if (m_Settings.IsUseGlobalResourceFile)
      {
        bool isPropertiesItem = (resourceFilePrjItem != null);

        if (isPropertiesItem)
        {
          prjItems            = resourceFilePrjItem.ProjectItems;
          resourceFilePrjItem = null;
          resourceFileDir     = System.IO.Path.Combine(resourceFileDir, resourceFileName); //append "Properties"/"My Project" because it exists
        } //if

        if (prjItems == null)
          return (null); //something went terribly wrong that never should have been possible

        if (string.IsNullOrEmpty(m_Settings.GlobalResourceFileName))
          resourceFileName = "Resources.resx"; //standard global resource file
        else
          resourceFileName = System.IO.Path.ChangeExtension(System.IO.Path.GetFileName(m_Settings.GlobalResourceFileName), "Resources.resx");

        try
        {
          //searches for the global resource
          resourceFilePrjItem = prjItems.Item(resourceFileName);
        }
        catch { }
      } //if

      if (resourceFilePrjItem == null)
      {
        #region not in project but file exists? -> ask user if delete
        string projectPath  = System.IO.Path.GetDirectoryName(m_Window.ProjectItem.ContainingProject.FullName),
               resourceFile = System.IO.Path.Combine(resourceFileDir, resourceFileName),
               designerFile = System.IO.Path.ChangeExtension(resourceFile, ".Designer." + m_ProjectExtension.Substring(0, 2)/*((m_IsCSharp) ? "cs" : "vb")*/);

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
          // Retrieve the path to the resource template.
          string itemPath = ((Solution2)m_Dte2.Solution).GetProjectItemTemplate("Resource.zip", m_ProjectExtension);

          //create a new project item based on the template
          /*prjItem =*/ prjItems.AddFromTemplate(itemPath, resourceFileName); //returns always null ...
          resourceFilePrjItem = prjItems.Item(resourceFileName);
        }
        catch (Exception ex)
        {
          Trace.WriteLine($"### OpenOrCreateResourceFile() - {ex.ToString()}");
          return (null);
        }
      } //if

      if (resourceFilePrjItem == null)
        return (null);

      //open the ResX file
      if (!resourceFilePrjItem.IsOpen[EnvDTEConstants.vsViewKindAny])
        resourceFilePrjItem.Open(EnvDTEConstants.vsViewKindDesigner);

      return (resourceFilePrjItem);
    }

    private void BuildAndUseResource(ProjectItem resourceFilePrjItem)
    {
      try
      {
        string resxFile = resourceFilePrjItem.Document.FullName;

        string nameSpace = m_Window.Project.Properties.Item("RootNamespace").Value.ToString();

        #region get namespace from ResX designer file
        foreach (ProjectItem item in resourceFilePrjItem.ProjectItems)
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

        //close the ResX file to modify it (force a checkout)
        resourceFilePrjItem.Document.Save(); //[13-07-10 DR]: MS has changed behavior of Close(vsSaveChanges.vsSaveChangesYes) in VS2012
        resourceFilePrjItem.Document.Close(vsSaveChanges.vsSaveChangesYes);

        string name,
               value,
               comment = string.Empty;
        StringResource stringResource = m_SelectedStringResource;// this.dataGrid1.CurrentItem as StringResource;

        name = stringResource.Name;
        value = stringResource.Text;

        if (!m_IsCSharp || (m_TextDocument.Selection.Text.Length > 0) && (m_TextDocument.Selection.Text[0] == '@'))
          value = value.Replace("\"\"", "\\\""); // [""] -> [\"] (change VB writing to CSharp)

        //add to the resource file (checking for duplicate)
        if (!AppendStringResource(resxFile, name, value, comment))
          return;

        //(re-)create the designer class
        VSLangProj.VSProjectItem vsPrjItem = resourceFilePrjItem.Object as VSLangProj.VSProjectItem;
        if (vsPrjItem != null)
          vsPrjItem.RunCustomTool();

        //get the length of the selected string literal and replace by resource call
        int replaceLength = m_TextDocument.Selection.Text.Length;
        if (replaceLength > 0)
        {
          string className    = System.IO.Path.GetFileNameWithoutExtension(resxFile).Replace('.', '_'),
                 aliasName    = className.Substring(0, className.Length - 6), //..._Resources -> ..._Res
                 resourceCall = string.Concat(aliasName, ".", name);

          bool isGlobalResourceFile        = m_Settings.IsUseGlobalResourceFile && string.IsNullOrEmpty(m_Settings.GlobalResourceFileName),
               isDontUseResourceUsingAlias = m_Settings.IsUseGlobalResourceFile && m_Settings.IsDontUseResourceAlias;

          if (isGlobalResourceFile)
          {
            //standard global resource file
            aliasName    = string.Concat("Glbl", aliasName);
            resourceCall = string.Concat("Glbl", resourceCall);
          } //if

          int oldRow = m_TextDocument.Selection.ActivePoint.Line;

          if (!isDontUseResourceUsingAlias)
          {
            //insert the resource using alias (if not yet)
            CheckAndAddAlias(nameSpace, className, aliasName);
          }
          else
          {
            //create a resource call like "Properties.SRB_Strings_Resources.myResText", "Properties.Resources.myResText", "Resources.myResText"
            int lastDotPos = nameSpace.LastIndexOf('.');
            string resxNameSpace = (lastDotPos >= 0) ? string.Concat(nameSpace.Substring(lastDotPos + 1), ".") : string.Empty;
            resourceCall = string.Concat(resxNameSpace, className, ".", name);
          } //else

          //insert the resource call, replacing the selected string literal
          m_TextDocument.Selection.Insert(resourceCall, (int)vsInsertFlags.vsInsertFlagsContainNewText);

          UpdateTableAndSelectNext(resourceCall.Length - replaceLength, m_TextDocument.Selection.ActivePoint.Line - oldRow);
        } //if
      }
      catch (Exception ex)
      {
        Trace.WriteLine($"### BuildAndUseResource() - {ex.ToString()}");
      }
    }

    private void CheckAndAddAlias(string nameSpace,
                                  string className,
                                  string aliasName)
    {
      string resourceAlias1 = $"using {aliasName} = ",
             resourceAlias2 = $"global::{nameSpace}.{className}";

      if (!m_IsCSharp)
      {
        resourceAlias1 = $"Imports {aliasName} = ";
        resourceAlias2 = $"{nameSpace}.{className}";
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

        string xmlPath = $"descendant::data[(attribute::name='{name}')]/descendant::value";
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
        Trace.WriteLine($"### AppendStringResource() - {ex.ToString()}");
        return (false);
      }

      return (true);
    }

    private void RemoveStringResource(int index)
    {
      m_StringResources.RemoveAt(index);
    }
    #endregion

    #region Ignore strings (Settings)
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
        string xml = System.IO.File.ReadAllText(m_SettingsFile, Encoding.UTF8);
        m_Settings = Settings.DeSerialize(xml);
      }
      else
        m_Settings = new Settings();  //defaults

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

      System.IO.File.WriteAllText(m_SettingsFile, m_Settings.Serialize(), Encoding.UTF8);

#if !IGNORE_METHOD_ARGUMENTS
      m_Settings.IgnoreMethodsArguments.Add("@@@disabled@@@");
#endif
    }
    #endregion

    #region Grid
    /// <summary>
    /// Find the nearest string literal to the actual cursor location in the table and select it.
    /// </summary>
    private void SelectNearestGridRow()
    {
      int curLine = m_TextDocument.Selection.ActivePoint.Line,
          curCol  = m_TextDocument.Selection.ActivePoint.LineCharOffset,
          rowNo   = -1,
          lineNo  = -1;

      List<StringResource> stringResources = GetAllStringResourcesInLine(curLine);

      if (stringResources.Count > 0)
      {
        #region current line

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

        #endregion
      }
      else
      {
        #region find nearest line (select only in grid)

        StringResource stringResourceBefore = GetLastStringResourceBeforeLine(curLine),
                       stringResourceBehind = GetFirstStringResourceBehindLine(curLine);

        if ((stringResourceBefore != null)
            && ((stringResourceBehind == null)
                || ((curLine - stringResourceBefore.Location.X) <= (stringResourceBehind.Location.X - curLine))))
        {
          rowNo  = m_StringResources.IndexOf(stringResourceBefore);
          lineNo = stringResourceBefore.Location.X;
          Debug.Print("> Found at ({0}, {1}): '{2}'", lineNo, stringResourceBefore.Location.Y, stringResourceBefore.Text);
        }
        else if (stringResourceBehind != null)
        {
          rowNo  = m_StringResources.IndexOf(stringResourceBehind);
          lineNo = stringResourceBehind.Location.X;
          Debug.Print("> Found at ({0}, {1}): '{2}'", lineNo, stringResourceBehind.Location.Y, stringResourceBehind.Text);
        } //else

        #endregion
      } //else

      if (rowNo >= 0)
      {
        this.SelectCell(rowNo, this.GetSelectedColIndex());

        if (lineNo == curLine)
          SelectStringInTextDocument();
      } //if
    }

    /// <summary>
    /// Updates the table and selects next entry.
    /// </summary>
    /// <param name="deltaLength">The length delta of the replacement (resource call length minus string literal length).</param>
    /// <param name="deltaLine">The line delta (when alias has been inserted).</param>
    private void UpdateTableAndSelectNext(int deltaLength,
                                          int deltaLine)
    {
      if ((deltaLength != 0) || (deltaLine != 0))
      {
        #region update entries

        int replacedLocationX = m_SelectedStringResource.Location.X,
            replacedLocationY = m_SelectedStringResource.Location.Y,
            startGridRowIndex = (deltaLine != 0) ? 0 : m_SelectedGridRowIndex + 1; //touch all when alias has been inserted
                                                                                   //else only the following entries on the same line

        for (int gridRowIndex = startGridRowIndex; gridRowIndex < m_StringResources.Count; ++gridRowIndex)
        {
          if (gridRowIndex == m_SelectedGridRowIndex)
            continue;

          StringResource stringResource = m_StringResources[gridRowIndex] as StringResource;

          if ((deltaLength != 0) && (stringResource.Location.X == replacedLocationX) && (stringResource.Location.Y > replacedLocationY))
          {
            //same line as replaced entry -> heed column shift and a possible alias insertion
            Debug.Print("{0}: col  {1} -> {2}, {3}", gridRowIndex,
                                                     stringResource.Location.ToString(),
                                                     stringResource.Location.X + deltaLine,
                                                     stringResource.Location.Y + deltaLength);

            stringResource.Offset(deltaLine, deltaLength);
          }
          else if (deltaLine != 0)
          {
            //heed alias insertion
            Debug.Print("{0}: line {1} -> {2}, {3}", gridRowIndex,
                                                     stringResource.Location.ToString(),
                                                     stringResource.Location.X + deltaLine,
                                                     stringResource.Location.Y + deltaLength);

            stringResource.Offset(deltaLine, 0);
          }
          else
            break; //nothing left to update
        } //for

#if never //[12-10-03 DR]: Indexing no longer in use
        //[12-10-03 DR]: search for resource names with selected name plus index ("_#")
        string name = string.Concat(m_SelectedStringResource.Name, "_");

        List<StringResource> stringResources = m_StringResources.FindAll(sr => sr.Name.StartsWith(name));

        if (stringResources.Count > 0)
        {
          //[12-10-03 DR]: calculate new indexes (remove first index entirely)
          int nameLength = name.Length,
              firstIndex = -1,
              index      = 1,
              indexValue;

          foreach (StringResource item in stringResources)
          {
            if (int.TryParse(item.Name.Substring(nameLength), out indexValue))
            {
              if (firstIndex == -1)
              {
                firstIndex = indexValue;
                item.Name = item.Name.Substring(0, nameLength - 1); //remove entire index ("_#")
              }
              else
              {
                item.Name = item.Name.Substring(0, nameLength);         //remove old index number
                item.Name = string.Concat(item.Name, index.ToString()); //append new index number
                ++index;
              } //else
            } //if
          } //foreach
        } //if
#endif

        #endregion
      } //if

      #region remove entry and goto next entry

      //this.ClearGrid();

      RemoveStringResource(m_SelectedGridRowIndex);

      //this.SetGridItemsSource(m_StringResources);
      this.RefreshGrid();

      this.SelectCell(m_SelectedGridRowIndex, this.GetSelectedColIndex());

      #endregion

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

    public void DoBrowse()
    {
      if (!IsSourceCode(m_FocusedTextDocumentWindow))
      {
        ClearStringResources();
        this.ClearGrid();
      }
      else
        BrowseForStrings(m_FocusedTextDocumentWindow, m_TextDocument.StartPoint, m_TextDocument.EndPoint);

      ResumeTextMove();
    }

    public void DoBrowse(TextPoint startPoint,
                         TextPoint endPoint)
    {
      SuspendTextMove();

      if (!IsSourceCode(m_FocusedTextDocumentWindow))
      {
        ClearStringResources();
        this.ClearGrid();
      }
      else
        BrowseForStrings(m_FocusedTextDocumentWindow, startPoint, endPoint);

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
      m_SelectedGridRowIndex       = this.GetSelectedRowIndex();

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

      ProjectItem resourceFilePrjItem = OpenOrCreateResourceFile();

      if (resourceFilePrjItem == null)
        return;

      BuildAndUseResource(resourceFilePrjItem);

      //SetEnabled();

      //MarkStringInTextDocument();

      m_LastDocumentLength = m_TextDocument.EndPoint.Line;
    }

    public void SaveSettings()
    {
      SaveIgnoreStrings();
    }

    #endregion //Public methods --------------------------------------------------------------------
  } //class
} //namespace
