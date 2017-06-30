# Revision History

### V1.6 R1 (22) [2017-06-30]

* **Fixed:** Issue with expression bodied properties no longer being parsed.
* **Fixed:** Missing standard resources for Help/About dialog of Visual Studio.
* **New:** Provide revision history as release notes in "Extensions and Updates".
* **New:** Option not to use the resource class using alias for global resources.  Thanks to **isaksavo**.

### V1.5 R3 (21) [2017-05-21]

* **Fixed:** Crash when parsing expression bodied properties due to a bug seen in VS2017 V15.2 (26430.6 +/-). But now, these property getters are ignored as long as the EnvDTE is not fixed to give proper start and end points.
* **Changed:** NuGet to use project.json.

### V1.5 R2 (20) [2016-12-10]

* **Changed:** Manifest changed to support only VS2015 and later from now on.
* **Changed:** No code generation anymore as VSPackage Builder runs only in VS2010.
* **Updated:** NuGet Packages.
* **New:** Using Visual Studio 2017 RC.
* **New:** extension.vsixmanifest updated to Visual Studio 2017 (V15.0).
* **New:** VSIX and DLL signed now.

### V1.5 R1 (19) [2016-02-27]

**Because this version has been built with VS2015, it is not clear, whether the extension runs in versions VS2010, VS2012 or VS2013** (though VS2013 seems to work fine when VS2015 is installed as well).

* **Fixed:** Handling new C#6 features:
	* Expression bodied properties, functions and indexers.
	* Interpolated strings ($"...") -> ignored completely.
* **Fixed:** Selecting the resources grid when unfocused was jumping weirdly.
* **Changed:** .NET Framework V4.5.2.
* **New:** Using Visual Studio 2015 (using result of VSPackage Builder without the extension as it sadly exists only for VS2010).
* **New:** Uses NuGet Packages:
	* Microsoft.VisualStudio.Shell.14.0
	* Microsoft.VisualStudio.Shell.Interop.12.0
	
### V1.3 R7 (18) [2015-08-07]

* **New:** extension.vsixmanifest updated for Visual Studio 2015 (V14.0)

### V1.3 R6 (17) [2014-01-12]

* **New:** extension.vsixmanifest updated for Visual Studio 2013 (V12.0)
* **New:** Option to use verbatim string literals (preceded by @ character) as "ignore" filter criteria.  Thanks to **Jean-Yves G.** for the suggestion.

### V1.3 R5 (16) [2013-07-10]

* **Fixed:** _VS2012:_ Broken behavior for automatically checking out ResX.  Thanks to **Piggy**.
* **Fixed:** _VS2012:_ Make selection background blue even when DataGrid is unfocused.

### V1.3 R4 (15) [2012-12-30]

* **Fixed:** Wrong column number parsed in single line function (e.g. get/set).  Thanks to **Game.Dev**.
* **Fixed:** Global resource file for VB did not work (should be in directory "My Project" instead of "Properties").

### V1.3 R3 (14) [2012-10-03]

* **Fixed:** Undo did incorrectly change the line numbers in table.
* **Changed:** Removed automatic resource name indexing.
* **Changed:** Ignore number strings now also ignores decimal numbers (" 1234.56 ").
* **Improved:** Some more information on the settings window.

### V1.3 R2 (13) [2012-09-14]

* **Fixed:** "Making" a string resource left the table of string literals blank (due to "improvement" in Build 11).
* **Improved:** Sped up inserting resource call when "making" a string resource.

### V1.3 R1 (12) [2012-09-12]

* **New:** Option to store all string resources in one resource file global to the project.

### V1.2 R4 (11) [2012-07-21]

* **Improved:** Rescan while editing no longer with timer and complete but direct and selective (only what has been changed) -> partly more responsive.
* **Reworked:** Code generators (VSPackage Builder) cleaned up (spaces no tabs, result code formatted to my likes).

### V1.2 R3 (9) [2012-06-27]

* **Fixed:** Handling empty string literals (locations behind did not match).
* **Fixed:** Closing and reopening tool window with same editor window active did not rescan.
* **Fixed:** Rescan did not look for nearest string resource (mark in table and mark next string literal when cursor in same line).
* **Fixed:** After editing the next string literal in same line has been marked.

### V1.2 R2 (7) [2012-06-11]

* **New:** extension.vsixmanifest updated for Visual Studio 2012 (V11.0)

### V1.2 R1 (6) [2012-06-09]

* **New:** Initial release for Visual Studio 2010

