# About
The String Resource Builder is a tool to support easy localization of code.  It is created for Visual Studio 2010 to 2015 to extract string literals from source code (C# and VB.NET) and put them into resource files (ResX), replacing the string literal by its resource call (also an using alias will be created).

All string literals are listed in a tool window with a suggested placeholder for the resource created from its text (see **Hints** below). The placeholder may be edited beforehand. Duplicates of placeholders will be recognized and the user will be questioned if he wants to use the already existing resource.

Per default for each source file one resource file will be created, thus it will be easy to share forms or classes between different solutions including their localizations. A resource file is named with the original source file name whose extension is replaced by ".Resources.resx" (e.g. "frmMain.cs" -> "frmMain.Resources.resx" plus its designer file, in VB.NET likewise).
There is also an option to have a single project resource file for string resources (since build 12).

## Hints
* This tool is not tested with character sets other than used in English and German (meaning non-ASCII), thus one should first **try it with a backup copy** of a corresponding solution (for localization I recommend programming in English as neutral language because a professional translator it will be more easy).
* When the tool window is open, editing the source will result in a refresh of the listed string literals (re-parsing of the code), which is automatically done after a pause of two seconds after the last change to the source code, which may lead to unexpected bahavior using the mouse to position the caret or copy&paste or move per drag&drop.  It is recommended to close the String Resource Builder tool window if not used or while coding (especially for big source files).

[Next: Usage](Usage.md)

