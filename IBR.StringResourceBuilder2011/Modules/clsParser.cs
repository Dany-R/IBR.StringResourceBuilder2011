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
                   Action<bool>         completedCallback)
    {
      m_Window          = window;
      m_StringResources = stringResources;
      m_Settings        = settings;
      m_DoProgress      = progressCallback;
      m_DoCompleted     = completedCallback;
      m_IsCSharp        = window.Caption.EndsWith(".cs");
    }

    #endregion //Constructor -----------------------------------------------------------------------

    #region Types
    #endregion //Types -----------------------------------------------------------------------------

    #region Fields

    private Window m_Window;
    private List<StringResource> m_StringResources;
    private Settings m_Settings;
    private Action<int> m_DoProgress;
    private Action<bool> m_DoCompleted;
    private bool m_IsCSharp;

    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties
    #endregion //Properties ------------------------------------------------------------------------

    #region Events
    #endregion //Events ----------------------------------------------------------------------------

    #region Private methods

    /// <summary>Parses for strings by iterating through the FileCodeModel.</summary>
    /// <param name="startPoint">The start point.</param>
    /// <param name="endPoint">The end point.</param>
    /// <param name="lastDocumentLength">Last length of the document.</param>
    private void ParseForStrings(TextPoint startPoint,
                                 TextPoint endPoint,
                                 int lastDocumentLength)
    {
      //0.35-0.06 seconds (threaded: 2.47-1.77 seconds)
      List<StringResource> stringResources = new List<StringResource>();

      bool isFullDocument          = startPoint.AtStartOfDocument && endPoint.AtEndOfDocument,
           isTextWithStringLiteral = true;
      int startLine                = startPoint.Line,
          startCol                 = startPoint.LineCharOffset,
          endLine                  = endPoint.Line,
          endCol                   = endPoint.LineCharOffset,
          documentLength           = endPoint.Parent.EndPoint.Line,
          insertIndex              = 0;

      if (isFullDocument)
        m_StringResources.Clear();
      else
      {
        #region document manipulated -> adapt string resources and locations

        //determine whether the text between startLine and endLine (including) contains double quotes
        EditPoint editPoint = startPoint.CreateEditPoint() as EditPoint2;
        if (!startPoint.AtStartOfLine)
          editPoint.StartOfLine();
        isTextWithStringLiteral = editPoint.GetLines(startLine, endLine + 1).Contains("\"");

        //move trailing locations behind changed lines if needed and
        //remove string resources on changed lines

        int lineOffset = documentLength - lastDocumentLength;
#if DEBUG_OUTPUT
        System.Diagnostics.Debug.Print("  Line offset is {0}", lineOffset);
#endif

        for (int i = m_StringResources.Count - 1; i >= 0; --i)
        {
          StringResource stringResource = m_StringResources[i];
          int            lineNo         = stringResource.Location.X;

          if (lineNo + lineOffset > endLine)
          {
            if (lineOffset != 0)
            {
#if DEBUG_OUTPUT
              System.Diagnostics.Debug.Print("  Move string literal from line {0} to {1}", lineNo, lineNo + lineOffset);
#endif
              stringResource.Offset(lineOffset, 0); //move
            } //if
          }
          else if (lineNo >= startLine)
          {
#if DEBUG_OUTPUT
            System.Diagnostics.Debug.Print("  Remove string literal {0} ({1}): {2}", i, stringResource.Location, stringResource.Text);
#endif
            m_StringResources.RemoveAt(i); //remove changed line
          }
          else if (insertIndex == 0)
          {
#if DEBUG_OUTPUT
            System.Diagnostics.Debug.Print("  List insert index is {0} / {1}", i + 1, m_StringResources.Count - 1);
#endif
            insertIndex = i + 1;
          } //else if
        } //for

        #endregion
      } //else

#if DEBUG_OUTPUT
      System.Diagnostics.Debug.Print("  Text has{0} string literals.", isTextWithStringLiteral ? string.Empty : " no");
#endif

      if (isTextWithStringLiteral)
      {
        CodeElements elements = m_Window.Document.ProjectItem.FileCodeModel.CodeElements;

        foreach (CodeElement element in elements)
        {
          ParseForStrings(element, m_DoProgress, stringResources, m_Settings, m_IsCSharp, startLine, endLine);

#if DEBUG
          if (element.Kind == vsCMElement.vsCMElementProperty)
          {
            CodeProperty prop = element as CodeProperty;

            if ((prop.Getter == null) && (prop.Setter == null))
            {
              //here we have an expression bodied property
              //if (m_IVsTextView != null)
              //{
              //  m_IVsTextView.
              //}
            }
          }
#endif

        } //foreach

#if DEBUG_OUTPUT
        System.Diagnostics.Debug.Print("  Found {0} string literals", stringResources.Count);
#endif

        if (isFullDocument)
          m_StringResources.AddRange(stringResources);
        else if (stringResources.Count > 0)
          m_StringResources.InsertRange(insertIndex, stringResources);
      } //if

      m_DoCompleted(isFullDocument || (stringResources.Count > 0));
    }

    /// <summary>Parses for strings by evaluating the element kind.</summary>
    /// <param name="element">The element.</param>
    /// <param name="progressWorker">The progress worker.</param>
    /// <param name="stringResources">The string resources.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="isCSharp">If set to <c>true</c> it is CSharp code.</param>
    /// <param name="startLine">The start line.</param>
    /// <param name="endLine">The end line.</param>
    private static void ParseForStrings(CodeElement element,
                                        Action<int> progressWorker,
                                        List<StringResource> stringResources,
                                        Settings settings,
                                        bool isCSharp,
                                        int startLine,
                                        int endLine)
    {
      if (element == null)
        return;

      try
      {
        if (element.StartPoint.Line > endLine)
          return;
      }
      catch (Exception ex)
      {
        //element.StartPoint not implemented in VS2017 15.2 (26430.6) for expression bodied property getters (before no getter element was available)
        System.Diagnostics.Debug.Print("### Error: ParseForStrings({0}): element.StartPoint.Line > endLine? {1} - {2}", element.Kind, ex.GetType().Name, ex.Message);
        return;
      }

      try
      {
        if (element.EndPoint.Line < startLine)
          return;
      }
      catch (Exception ex)
      {
        //element.EndPoint invalid when deleting or cutting text
        System.Diagnostics.Debug.Print("### Error: ParseForStrings(): element.EndPoint < startLine? {0} - {1}", ex.GetType().Name, ex.Message);
        return;
      }

#if DEBUG_OUTPUT
      string elementName = string.Empty;
      if (!isCSharp || (element.Kind != vsCMElement.vsCMElementImportStmt))
        try { elementName = element.Name; } catch {}
      System.Diagnostics.Debug.Print("  > ParseForStrings({0} '{1}')", element.Kind, elementName);
#endif

      if (element.IsCodeType
          && ((element.Kind == vsCMElement.vsCMElementClass)
              || (element.Kind == vsCMElement.vsCMElementStruct)))
      {
        CodeElements elements = (element as CodeType).Members;

        foreach (CodeElement element2 in elements)
          ParseForStrings(element2, progressWorker, stringResources, settings, isCSharp, startLine, endLine);
      }
      else if (element.Kind == vsCMElement.vsCMElementNamespace)
      {
        CodeElements elements = (element as CodeNamespace).Members;

        foreach (CodeElement element2 in elements)
          ParseForStrings(element2, progressWorker, stringResources, settings, isCSharp, startLine, endLine);
      }
      else if (element.Kind == vsCMElement.vsCMElementProperty)
      {
        CodeProperty prop = element as CodeProperty;

        //CodeElement.StartPoint not implemented in VS2017 15.2 (26430.6) for expression bodied property getters
        //because before the expression bodied properties had Getter and Setter == null
        bool getterHasStartPoint = (prop.Getter != null) && (prop.Getter as CodeElement).HasStartPoint(),
             setterHasStartPoint = (prop.Setter != null) && (prop.Setter as CodeElement).HasStartPoint();

        if (getterHasStartPoint)
          ParseForStrings(prop.Getter as CodeElement, progressWorker, stringResources, settings, isCSharp, startLine, endLine);

        if (setterHasStartPoint)
          ParseForStrings(prop.Setter as CodeElement, progressWorker, stringResources, settings, isCSharp, startLine, endLine);

        if (!getterHasStartPoint && !setterHasStartPoint)
        {
          //expression bodied property
          int lineNo = ParseForStrings(element, stringResources, settings, isCSharp, startLine, endLine);
          if (lineNo > -1)
            progressWorker(lineNo);
        } //if
      }
      else if ((element.Kind == vsCMElement.vsCMElementFunction)
               || (element.Kind == vsCMElement.vsCMElementVariable))
      {
        if ((element.Kind == vsCMElement.vsCMElementFunction) && settings.IgnoreMethod(element.Name))
          return;

        int lineNo = ParseForStrings(element, stringResources, settings, isCSharp, startLine, endLine);
        if (lineNo > -1)
          progressWorker(lineNo);
      } //else if
    }

    /// <summary>Parses for strings by iterating through the text lines.</summary>
    /// <param name="element">The element.</param>
    /// <param name="stringResources">The string resources.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="isCSharp">If set to <c>true</c> it is CSharp code.</param>
    /// <param name="startLine">The start line.</param>
    /// <param name="endLine">The end line.</param>
    /// <returns>The last parsed line number or -1.</returns>
    private static int ParseForStrings(CodeElement element,
                                       List<StringResource> stringResources,
                                       Settings settings,
                                       bool isCSharp,
                                       int startLine,
                                       int endLine)
    {
      TextPoint  startPoint = element.StartPoint,
                 endPoint   = element.EndPoint;
      EditPoint2 editPoint  = null;

      try
      {
        if (element.Kind == vsCMElement.vsCMElementFunction)
        {
#if DEBUG_OUTPUT
          System.Diagnostics.Debug.Print("    function kind: {0}", (element as CodeFunction).FunctionKind);
#endif

          try
          {
            //we want to have the body only (throws COMException when inspecting an expression bodied function)
            startPoint = element.GetStartPoint(vsCMPart.vsCMPartBody);
            endPoint   = element.GetEndPoint(vsCMPart.vsCMPartBody);
          }
          catch (Exception ex)
          {
            System.Diagnostics.Debug.Print("ParseForStrings(vsCMElementFunction, line {0}): {1} - {2}", startPoint.Line, ex.GetType().Name, ex.Message);
          }
        } //if

        editPoint = startPoint.CreateEditPoint() as EditPoint2;

        //if ((element.Kind == vsCMElement.vsCMElementVariable) && (startPoint.LineCharOffset > 1))
        //  editPoint.CharLeft(startPoint.LineCharOffset - 1);

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
        //int endPoint      = element.EndPoint.Line,
        //    endColumn     = element.EndPoint.LineCharOffset,
        //    absoluteEnd   = element.EndPoint.AbsoluteCharOffset,
        //    editLine      = editPoint.Line,
        //    editColumn    = editPoint.LineCharOffset,
        //    absoluteStart = editPoint.AbsoluteCharOffset,
        //    editLength    = (editLine == endPoint) ? (absoluteEnd - absoluteStart + 1)
        //                                           : (editPoint.LineLength - editColumn + 1);

        //while ((editLine < endPoint) || ((editLine == endPoint) && (editColumn <= endColumn)))
        //{
        //  string textLine = editPoint.GetText(editLength);

        //  //System.Diagnostics.Debug.Print(">>>{0}<<<", textLine);

        //  if (!string.IsNullOrEmpty(textLine.Trim()))
        //    ParseForStrings(textLine, editLine, editColumn, stringResources, settings);

        //  editPoint.LineDown(1);
        //  editPoint.StartOfLine();

        //  editLine      = editPoint.Line;
        //  editColumn    = editPoint.LineCharOffset;
        //  absoluteStart = editPoint.AbsoluteCharOffset;
        //  editLength    = (editLine == endPoint) ? (absoluteEnd - absoluteStart + 1)
        //                                         : (editPoint.LineLength - editColumn + 1);
        //} //while
        #endregion

        //this is much faster (by factors)!!!
        int      editLine                = editPoint.Line,
                 editColumn              = editPoint.LineCharOffset;
        string   text                    = editPoint.GetText(endPoint);
        string[] txtLines                = text.Replace("\r", string.Empty).Split('\n');
        bool     isComment               = false;
#if IGNORE_METHOD_ARGUMENTS
        bool     isIgnoreMethodArguments = false;
#endif

        foreach (string txtLine in txtLines)
        {
          if ((editLine >= startLine) && (editLine <= endLine))
          {
            //this is a changed text line in the block

#if !IGNORE_METHOD_ARGUMENTS
            if (txtLine.Contains("\""))
              ParseForStrings(txtLine, editLine, editColumn, stringResources, settings, isCSharp, ref isComment);
#else
            if (line.Contains("\""))
              ParseForStrings(line, editLine, editColumn, stringResources, settings, ref isComment, ref isIgnoreMethodArguments);
#endif
          } //if

          ++editLine;

          //only for the first line of the text block LineCharOffset will be used
          if (editColumn > 1)
            editColumn = 1;
        } //foreach
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.Print("### Error: ParseForStrings(): {0} - {1}", ex.GetType().Name, ex.Message);
      }

      return (endPoint?.Line??(-1));
    }

    /// <summary>Parses for strings by iterating through the parts between double quotes.</summary>
    /// <param name="txtLine">The text line.</param>
    /// <param name="lineNo">The line number.</param>
    /// <param name="colNo">The column number.</param>
    /// <param name="stringResources">The string resources.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="isCSharp">If set to <c>true</c> it is CSharp code.</param>
    /// <param name="isComment">If set to <c>true</c> it starts (in) or ends (out) with a comment.</param>
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
           isAtString           = false, // @"..."
           isInterpolatedString = false; // $"..."
      StringBuilder txt = new StringBuilder();
      int pos = 0;

      string[] parts = txtLine.Split('"');

      for (int i = 0; i < parts.Length; ++i)
      {
        string part = parts[i];

        bool isEmptyPart = (part.Length == 0);

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

          isAtString           = !isCSharp || part.EndsWith("@");
          isInterpolatedString = isCSharp && (part.EndsWith("$") || part.EndsWith("$@"));

          stringPos.Add(pos);

          isInString = true;
          #endregion
        }
        else //if (isInString)
        {
          if (isInterpolatedString)
          {
            //[16-02-20 DR]: ignored for now
            #region $-string ($"...") or $@-string ($@"...")
            //e.g. $"{m_MyString4} is \"also\" a {m_MyString3 + " value."} {{x}} {{\"x\"}}"
            //     $@"{m_MyString4} is ""also"" a {m_MyString3 + " value."} {{x}} {{""x""}}"
            //I hope that nobody nests interpolated strings as in $"foo {bar + $"some {bar2} value."}"

#if DEBUG
            //string a = "1",
            //       b = "2",
            //       x = $"{a} is \"also\" a {b + " \"value\"."} {{x}} {{\"x\"}}"
            //         + $@"{a} is ""also"" a {b + @" ""value""."
            //                                   + "substring "} {{x}} {{""x""}}"
            //         + $"foo {a + $"some {b} value."}"
            //         + "! {a}";
#endif

            //SkipInterpolatedString
            if (isAtString && isEmptyPart)
            {
              //handle leading double quotes
              isInString = SkipDoubleQuotes(txtLine, ref pos, ref i);
              continue;
            } //if

            if (!isAtString || !isEmptyPart)
              pos += part.Length + 1;

            if (part.Contains("{{"))
              part = part.Replace("{{", "##");
            if (part.Contains("}}"))
              part = part.Replace("}}", "##");
            if (part.Contains(@"\\"))
              part = part.Replace(@"\\", "##");

            int openCurlyBraces = IterateCurlyBraceBlocks(part);

            if (openCurlyBraces > 0)
            {
              #region skip nested string

              //LanguageServices MatchBraces BeginParse

              //if (part.EndsWith("$") || part.EndsWith("$@"))
              //{
              //  //SkipInterpolatedString
              //}
              //else
              ////SkipNestedString
              //if (part.EndsWith("@"))
              //{
              //  SkipDoubleQuotes(txtLine, ref pos, ref i);
              //}
              //else if (part.EndsWith("\\"))
              //{ }
#if DEBUG
              System.Diagnostics.Debug.Print("ignoring >>{0}<<", txtLine);
              System.Diagnostics.Debugger.Break();
#endif
              break; //[16-02-27 DR]: ignored completely

              #endregion
            }

            if (isAtString)
              isInString = !SkipDoubleQuotes(txtLine, ref pos, ref i); //handle trailing double quotes (inverted result)
            else if (!part.EndsWith(@"\"))
              isInString = false;
            #endregion
          }
          else if (isAtString)
          {
            #region @-string (@"...") or Basic string
            if (!isEmptyPart)
            {
              if (!settings.IsIgnoreVerbatimStrings)
                txt.Append(part);
              pos += part.Length;
            } //if

            //here double quotes are given as pairs
            int count = CountDoubleQuotes(txtLine, pos);

            if (count > 1)
            {
              //skip empty parts from double quote pairs
              i += count - 1;

              while (count >= 2)
              {
                if (!settings.IsIgnoreVerbatimStrings)
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

          if (!isInString && (!isAtString || !settings.IsIgnoreVerbatimStrings))
            txt.Append("\0"); //for later splitting
        } //else
      } //for

      if (txt.Length > 0)
      {
        #region filter and put in stringResources
        if (txt.ToString().EndsWith("\0"))
          txt.Remove(txt.Length - 1, 1);

        parts = txt.ToString().Split('\0');
        for (int i = 0; i < parts.Length; ++i)
        {
          string draftName = parts[i],
                 name      = string.Empty;

          if (settings.IgnoreString(draftName))
            continue;

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

          //[12-10-03 DR]: Indexing no longer in use
          //int count = GetFirstFreeNameIndex(name, stringResources);
          //if (count > 0)
          //  name += "_" + count.ToString();

          stringResources.Add(new StringResource(name, draftName, new System.Drawing.Point(lineNo, stringPos[i] + colNo - 1)));
        } //for
        #endregion
      } //if
    }

    /// <summary>Skips the double quotes.</summary>
    /// <param name="txtLine">The text line.</param>
    /// <param name="pos">The position.</param>
    /// <param name="i">The i.</param>
    /// <returns>
    /// <c>true</c> when the number of found double qoutes is even, meaning we are still within a string;
    /// otherwise <c>false</c>.
    /// </returns>
    private static bool SkipDoubleQuotes(string txtLine,
                                         ref int pos,
                                         ref int i)
    {
      //skip empty parts from double quote pairs:
      //"x"     = true(0),  -> never occurs here (called only when part is empty)
      //""      = false(1),
      //"""x"   = true(2),
      //""""    = false(3),
      //"""""x" = true(4)

      int count = CountDoubleQuotes(txtLine, pos);

      if (count > 0)
      {
        i += count - 1;

        pos += count;
      } //if

      return ((count & 1) == 0); //isInString
    }

    /// <summary>Counts the consecutive double quotes.</summary>
    /// <param name="txtLine">The text line.</param>
    /// <param name="startPos">The start position.</param>
    /// <returns>The number of consecutive double quotes.</returns>
    private static int CountDoubleQuotes(string txtLine,
                                         int startPos)
    {
      int count = startPos;

      while ((count < txtLine.Length) && (txtLine[count] == '"'))
        ++count;

      count -= startPos;

      return (count);
    }

    /// <summary>Iterates through the curly brace blocks.</summary>
    /// <param name="part">The part.</param>
    /// <returns>The number of unclosed (still open) curly brace blocks.</returns>
    private static int IterateCurlyBraceBlocks(string part)
    {
      int openCurlyBraces = 0;

      for (int j = 0; j < part.Length; ++j)
      {
        char c = part[j];

        if (c == '{')
          ++openCurlyBraces;
        else if (c == '}')
          --openCurlyBraces;
      } //for

      return (openCurlyBraces);
    }

    //[12-10-03 DR]: Indexing no longer in use
    //private static int GetFirstFreeNameIndex(string name,
    //                                         List<StringResource> stringResources)
    //{
    //  System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("^" + name + @"(_\d+)?$");

    //  List<StringResource> names = stringResources.FindAll(sr => regex.IsMatch(sr.Name));

    //  if ((names.Count == 0) || !names[0].Name.Equals(name))
    //    return (0);

    //  for (int i = 1; i < names.Count; ++i)
    //  {
    //    if (!names[i].Name.EndsWith("_" + i.ToString()))
    //      return (i);
    //  } //for

    //  return (names.Count);
    //}

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
                                       Action<bool> completedCallback,
                                       TextPoint startPoint,
                                       TextPoint endPoint,
                                       int lastDocumentLength)
    {
      Parser parser = new Parser(window, stringResources, settings, progressCallback, completedCallback);
      parser.ParseForStrings(startPoint, endPoint, lastDocumentLength);
    }

    #endregion //Public methods --------------------------------------------------------------------
  } //class
} //namespace
