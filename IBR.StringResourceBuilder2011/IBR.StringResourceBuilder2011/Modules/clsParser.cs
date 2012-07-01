//#define USE_THREAD
//#define IGNORE_METHOD_ARGUMENTS

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using EnvDTE80;



namespace IBR.StringResourceBuilder2011.Modules
{
  class Parser
  {
    #region Constructor

    private Parser(Window               window,
                   List<StringResource> stringResources,
                   Settings             settings,
                   Action<int>          progressCallback,
                   Action               completedCallback)
    {
      m_Window          = window;
      m_StringResources = stringResources;
      m_Settings        = settings;
      m_DoProgress      = progressCallback;
      m_DoCompleted     = completedCallback;
      m_IsCSharp        = window.Caption.EndsWith(".cs");

#if USE_THREAD
      m_Worker.WorkerReportsProgress      = true;
      m_Worker.WorkerSupportsCancellation = false;

      m_Worker.DoWork += m_Worker_DoWork;

      if (progressCallback != null)
        m_Worker.ProgressChanged += m_Worker_ProgressChanged;

      if (completedCallback != null)
        m_Worker.RunWorkerCompleted += m_Worker_RunWorkerCompleted;
#endif
    }

    #endregion //Constructor -----------------------------------------------------------------------

    #region Types

#if USE_THREAD
    private struct WorkerArgs
    {
      public WorkerArgs(Window window,
                        Settings settings)
      {
        Window   = window;
        Settings = settings;
      }

      public Window   Window;
      public Settings Settings;
    }
#endif

    #endregion //Types -----------------------------------------------------------------------------

    #region Fields

    private Window m_Window;
    private List<StringResource> m_StringResources;
    private Settings m_Settings;
    private Action<int> m_DoProgress;
    private Action m_DoCompleted;
    private bool m_IsCSharp;

#if USE_THREAD
    BackgroundWorker m_Worker = new BackgroundWorker();
#endif

    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties
    #endregion //Properties ------------------------------------------------------------------------

    #region Events
    #endregion //Events ----------------------------------------------------------------------------

    #region Private methods

    private void ParseForStrings()
    {
#if USE_THREAD
      //2.47-1.77 seconds
      WorkerArgs arguments = new WorkerArgs(m_Window, m_Settings);
      m_Worker.RunWorkerAsync(arguments);
#else
      //0.35-0.06 seconds
      List<StringResource> stringResources = new List<StringResource>();
      m_StringResources.Clear();

      CodeElements elements = m_Window.Document.ProjectItem.FileCodeModel.CodeElements;

      foreach (CodeElement element in elements)
      {
        ParseForStrings(element, m_DoProgress, stringResources, m_Settings, m_IsCSharp);
      } //foreach

      m_StringResources.AddRange(stringResources);
      m_DoCompleted();
#endif
    }

#if USE_THREAD
    private void m_Worker_DoWork(object sender, DoWorkEventArgs e)
    {
      BackgroundWorker worker = sender as BackgroundWorker;
      WorkerArgs arguments = (WorkerArgs)e.Argument;

      List<StringResource> stringResources = new List<StringResource>();

      CodeElements elements = arguments.Window.Document.ProjectItem.FileCodeModel.CodeElements;

      foreach (CodeElement element in elements)
      {
        if (worker.CancellationPending)
        {
          e.Cancel = true;
          break;
        } //if

        ParseForStrings(element, worker, stringResources, arguments.Settings, m_IsCSharp);
      } //foreach

      e.Result = stringResources;
    }

    private void m_Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
      m_DoProgress(e.ProgressPercentage);
    }

    private void m_Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      m_StringResources.Clear();

      if (!e.Cancelled && (e.Error == null))
        m_StringResources.AddRange(e.Result as List<StringResource>);

      m_DoCompleted();
    }
#endif

    private static void ParseForStrings(CodeElement element,
#if USE_THREAD
                                        BackgroundWorker worker,
#else
                                        Action<int> worker,
#endif
                                        List<StringResource> stringResources,
                                        Settings settings,
                                        bool isCSharp)
    {
#if USE_THREAD
      if (worker.CancellationPending)
        return;
#endif

      if (element == null)
        return;

#if DEBUG
      string elementName = string.Empty;
      if (!isCSharp || (element.Kind != vsCMElement.vsCMElementImportStmt))
        try { elementName = element.Name; } catch {}
      System.Diagnostics.Debug.Print("ParseForStrings({0} '{1}')", element.Kind, elementName);
#endif

      EditPoint2 editPoint  = null;
      TextPoint  startPoint = null,
                 endPoint   = null;

      if (element.IsCodeType
          && ((element.Kind == vsCMElement.vsCMElementClass)
              || (element.Kind == vsCMElement.vsCMElementStruct)))
      {
        CodeElements elements = (element as CodeType).Members;

        foreach (CodeElement element2 in elements)
          ParseForStrings(element2, worker, stringResources, settings, isCSharp);
      }
      else if (element.Kind == vsCMElement.vsCMElementNamespace)
      {
        CodeElements elements = (element as CodeNamespace).Members;

        foreach (CodeElement element2 in elements)
          ParseForStrings(element2, worker, stringResources, settings, isCSharp);
      }
      else if (element.Kind == vsCMElement.vsCMElementProperty)
      {
        CodeProperty prop = element as CodeProperty;
        if (prop.Getter != null)
          ParseForStrings(prop.Getter as CodeElement, worker, stringResources, settings, isCSharp);
        if (prop.Setter != null)
          ParseForStrings(prop.Setter as CodeElement, worker, stringResources, settings, isCSharp);
      }
      else if ((element.Kind == vsCMElement.vsCMElementFunction)
               || (element.Kind == vsCMElement.vsCMElementVariable))
      {
        if ((element.Kind == vsCMElement.vsCMElementFunction) && settings.IgnoreMethod(element.Name))
          return;

        try
        {
          if (element.Kind == vsCMElement.vsCMElementFunction)
          {
            startPoint = element.GetStartPoint(vsCMPart.vsCMPartBody);
            endPoint   = element.GetEndPoint(vsCMPart.vsCMPartBody);
          }
          else
          {
            startPoint = element.StartPoint;
            endPoint   = element.EndPoint;
          } //else

          editPoint = startPoint.CreateEditPoint() as EditPoint2;

          if ((element.Kind == vsCMElement.vsCMElementVariable) && (startPoint.LineCharOffset > 1))
            editPoint.CharLeft(startPoint.LineCharOffset - 1);

          //#if DEBUG
          //        if (element.Kind == vsCMElement.vsCMElementFunction)
          //          System.Diagnostics.Debug.Print("got a function, named: {0} in line {1} to {2}", element.Name, element.StartPoint.Line, element.EndPoint.Line);
          //        else
          //          System.Diagnostics.Debug.Print("got a variable, named: {0} in line {1} to {2}", element.Name, element.StartPoint.Line, element.EndPoint.Line);
          //#endif
          //if (element.Children.Count > 0)
          //  editPoint = element.Children.Item(element.Children.Count).EndPoint.CreateEditPoint() as EditPoint2;
          //else
          //  editPoint = element.StartPoint.CreateEditPoint() as EditPoint2;
          //#if DEBUG
          //        if (element.Children.Count > 0) System.Diagnostics.Debug.Print("      line {0} to {1}", editPoint.Line, element.EndPoint.Line);
          //#endif

          #region veeeeeery sloooooow
          //int endLine = element.EndPoint.Line,
          //    endColumn = element.EndPoint.LineCharOffset,
          //    absoluteEnd = element.EndPoint.AbsoluteCharOffset,
          //    editLine = editPoint.Line,
          //    editColumn = editPoint.LineCharOffset,
          //    absoluteStart = editPoint.AbsoluteCharOffset,
          //    editLength = (editLine == endLine) ? (absoluteEnd - absoluteStart + 1)
          //                                       : (editPoint.LineLength - editColumn + 1);

          //while ((editLine < endLine) || ((editLine == endLine) && (editColumn <= endColumn)))
          //{
#if USE_THREAD
          //  if (worker.CancellationPending)
          //    return; 
#endif

          //  string textLine = editPoint.GetText(editLength);

          //  //            System.Diagnostics.Debug.Print(">>>{0}<<<", textLine);

          //  if (!string.IsNullOrEmpty(textLine.Trim()))
          //    ParseForStrings(textLine, editLine, editColumn, stringResources, settings);

          //  editPoint.LineDown(1);
          //  editPoint.StartOfLine();

          //  editLine = editPoint.Line;
          //  editColumn = editPoint.LineCharOffset;
          //  absoluteStart = editPoint.AbsoluteCharOffset;
          //  editLength = (editLine == endLine) ? (absoluteEnd - absoluteStart + 1)
          //                                     : (editPoint.LineLength - editColumn + 1);
          //} //while
          #endregion

          //this is much faster (by factors)!!!
          int      editLine   = editPoint.Line,
                   editColumn = 1;// editPoint.LineCharOffset;
          string   text       = editPoint.GetText(endPoint);
          string[] lines      = text.Replace("\r", string.Empty).Split('\n');
          bool     isComment  = false;
#if IGNORE_METHOD_ARGUMENTS
          bool     isIgnoreMethodArguments = false;
#endif

          foreach (string line in lines)
          {
#if USE_THREAD
            if (worker.CancellationPending)
              return;
#endif

            //editColumn = line.Length - line.TrimStart(' ', '\t').Length;

#if !IGNORE_METHOD_ARGUMENTS
            if (/*!string.IsNullOrEmpty(line.Trim()) &&*/ line.Contains("\""))
              ParseForStrings(line, editLine, editColumn, stringResources, settings, isCSharp, ref isComment);
#else
            if (/*!string.IsNullOrEmpty(line.Trim()) &&*/ line.Contains("\""))
              ParseForStrings(line, editLine, editColumn, stringResources, settings, ref isComment, ref isIgnoreMethodArguments);
#endif

            ++editLine;
          } //foreach
        }
        catch (Exception ex)
        {
          System.Diagnostics.Debug.Print("Error in ParseForStrings: {0}", ex.Message);
        }
      } //else if

      if (endPoint != null)
      {
#if USE_THREAD
        worker.ReportProgress(endPoint.Line);
#else
        worker(endPoint.Line);
#endif
      } //if
    }

    private static void ParseForStrings(string txtLine,
                                        int lineNo,
                                        int colNo,
                                        List<StringResource> stringResources,
                                        Settings settings,
                                        bool isCSharp,
#if !IGNORE_METHOD_ARGUMENTS
                                        ref bool isComment
#else
                                        ref bool isComment,
                                        ref bool isIgnoreMethodArguments
#endif
                                       )
    {
      List<int> stringPos = new List<int>();

      bool isInString = false,
           isAtString = false; // @"..."
      StringBuilder txt = new StringBuilder();
      int pos = 0;

      string[] parts = txtLine.Split('"');

      for (int i = 0; i < parts.Length; ++i)
      {
        string part = parts[i];

        if (!isInString)
        {
          #region part is outside of strings

          #region handle comment
          if (isCSharp)
          {
            part = HandleComment(part, ref isComment);

            if (part.Contains("//")) //line comment -> ignore the rest of the parts
              break;
          }
          else
          {
            if (part.Contains("'")) //line comment -> ignore the rest of the parts
              break;

            //REM line comment -> ignore the rest of the parts
            // <whitespace|nonletter>REM<whitespace|nonletter>
            // <whitespace|nonletter>REM ccc
            int remPos = part.ToUpper().IndexOf("REM");
            if (remPos > -1)
            {
              char c = (remPos + 3 < part.Length) ? part[remPos + 3] : '\0';
              if ((c == '\0') || !(char.IsLetterOrDigit(c) || (c == '_')))
              {
                c = (remPos > 0) ? part[remPos - 1] : '\0';
                if (char.IsLetterOrDigit(c) || char.IsWhiteSpace(c))
                  break;
              } //if
            } //if
          } //else
          #endregion

          pos += part.Length + 1;

          if (isComment)
            continue;

          if (isCSharp && part.EndsWith("'")) //begin of char '"'
            continue;

#if IGNORE_METHOD_ARGUMENTS
          part = HandleIgnoreMethodArguments(part, ref isIgnoreMethodArguments);
          if (isIgnoreMethodArguments || string.IsNullOrEmpty(part.Trim()))
            continue;
#endif

          isAtString = isCSharp && part.EndsWith("@");

          stringPos.Add(pos);

          isInString = true;
          #endregion
        }
        else //if (isInString)
        {
          if (isAtString || !isCSharp)
          {
            #region @-string (@"...") or Basic string
            if (part != string.Empty)
            {
              txt.Append(part);
              pos += part.Length;
            } //if

            //here double quotes are given as pairs
            int count = pos;
            while ((count < txtLine.Length) && (txtLine[count] == '"'))
              ++count;
            count -= pos;

            if (count > 1)
            {
              //skip empty parts from double quote pairs
              i += count - 1;

              while (count >= 2)
              {
                txt.Append("\"\"");
                count -= 2;
                pos += 2;
              } //while
            } //if

            if (count == 1)
            {
              isInString = false;
              ++pos;
            } //if
            #endregion
          }
          else
          {
            #region normal string (CSharp only)
            //keep string as it is in the editor
            pos += part.Length + 1;

            //replace double backslash and undo it later to find single ending backslash
            if (part.Contains(@"\\"))
              part = part.Replace(@"\\", "\001");

            //single ending backslash escaping a double quote character
            if (part.EndsWith(@"\"))
              part = part/*.Remove(part.Length - 1)*/ + "\"";
            else
              isInString = false;

            //undo
            if (part.Contains("\001"))
              part = part.Replace("\001", @"\\");

            txt.Append(part);
            #endregion
          } //else

          if (!isInString)
            txt.Append("\0"); //for later splitting
        } //else
      } //for

      if (txt.Length > 0)
      {
        #region put in stringResources
        if (txt.ToString().EndsWith("\0"))
          txt.Remove(txt.Length - 1, 1);

        parts = txt.ToString().Split('\0');
        for (int i = 0; i < parts.Length; ++i)
        {
          string draftName = parts[i],
                 name = string.Empty;

          for (int c = 0; c < draftName.Length; ++c)
          {
            if (char.IsWhiteSpace(draftName[c]))
              name += '_';
            else if (char.IsLetterOrDigit(draftName[c]) || (draftName[c] == '_'))
              name += draftName[c];
          } //for

          if (name.Length == 0)
            name = stringResources.Count.ToString();

          if (char.IsDigit(name[0]))
            name = "_" + name;

          int count = GetFirstFreeNameIndex(name, stringResources);
          if (count > 0)
            name += "_" + count.ToString();

          if (settings.IgnoreString(draftName))
            continue;

          stringResources.Add(new StringResource(name, draftName, new System.Drawing.Point(lineNo, stringPos[i] + colNo - 1)));
        } //for
        #endregion
      } //if
    }

    private static int GetFirstFreeNameIndex(string name,
                                             List<StringResource> stringResources)
    {
      System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("^" + name + @"(_\d+)?$");

      List<StringResource> names = stringResources.FindAll(delegate(StringResource sr)
      {
        return (regex.IsMatch(sr.Name));
      });

      if ((names.Count == 0) || !names[0].Name.Equals(name))
        return (0);

      for (int i = 1; i < names.Count; ++i)
      {
        if (!names[i].Name.EndsWith("_" + i.ToString()))
          return (i);
      } //for

      return (names.Count);
    }

    private static string HandleComment(string txtPart,
                                        ref bool isComment)
    {
      int pos = -1;

      if (isComment)
      {
        //replace all characters of the comment by blank spaces
        pos = txtPart.IndexOf("*/");
        if (pos == -1)
          return (new string(' ', txtPart.Length)); //no end yet

        isComment = false;
        txtPart = new string(' ', pos + 2) + txtPart.Substring(pos + 2);
      } //if

      int pos2 = txtPart.IndexOf("//");
      pos = txtPart.IndexOf("/*");
      if ((pos == -1) || ((pos2 > -1) && (pos > pos2)))
        return (txtPart);

      isComment = true;
      txtPart = txtPart.Substring(0, pos) + HandleComment(txtPart.Substring(pos), ref isComment);
      return (txtPart);
    }

#if IGNORE_METHOD_ARGUMENTS
    private static string HandleIgnoreMethodArguments(string txtPart,
                                                      Settings settings,
                                                      ref bool isIgnoreMethodArguments,
                                                      ref int parentheses)
    {
      int pos = 0;
      int pos2 = 0;

      if (isIgnoreMethodArguments)
      {
        //ToDo: HandleIgnoreMethodArguments - VB.NET
        //count parentheses
        //bool isQuote = false,
        //     isChar = false,
        //     isEscape = false;
        while ((pos > 0) && (pos < txtPart.Length))
        {
          char c = txtPart[pos];
          txtPart = txtPart.Remove(pos).Insert(pos, " ");
          ++pos;

          if (c == '(')
            ++parentheses;
          else if (c == ')')
          {
            --parentheses;

            if (parentheses == 0)
            {
              isIgnoreMethodArguments = false;
              break;
            } //if
          } //else if
        } //while
      } //if

      foreach (string ignoreMethod in settings.IgnoreMethodsArguments)
      {
        pos = txtPart.LastIndexOf(ignoreMethod);
        if (pos < 0)
          continue;

        pos2 = pos - 1;
        if ((pos > 0) && (Char.IsLetterOrDigit(txtPart, pos2) || (txtPart[pos2] == '_')))
          continue;

        pos2 = pos + ignoreMethod.Length;
        if ((pos2 < txtPart.Length) && (Char.IsLetterOrDigit(txtPart, pos2) || (txtPart[pos2] == '_')))
          continue;

        isIgnoreMethodArguments = true;
        txtPart = txtPart.Substring(0, pos)
                  + new string(' ', ignoreMethod.Length)
                  + HandleIgnoreMethodArguments(txtPart.Substring(pos2), settings, ref isIgnoreMethodArguments, ref parentheses);
        break;
      } //foreach

      return (txtPart);
    }
#endif

    #endregion //Private methods -------------------------------------------------------------------

    #region Public methods

    public static void ParseForStrings(Window window,
                                       List<StringResource> stringResources,
                                       Settings settings,
                                       Action<int> progressCallback,
                                       Action completedCallback)
    {
      Parser parser = new Parser(window, stringResources, settings, progressCallback, completedCallback);
      parser.ParseForStrings();
    }

    #endregion //Public methods --------------------------------------------------------------------
  } //class
} //namespace
