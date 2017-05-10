using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security.Principal;
using System.Security.AccessControl;
using System.Runtime.InteropServices;



namespace IBR.StringResourceBuilder2011
{
  internal static class CUtil
  {
    #region Constructor
    #endregion //Constructor -----------------------------------------------------------------------

    #region Types

    [DllImport("shfolder.dll", CharSet = CharSet.Auto)]
    internal static extern int SHGetFolderPath(IntPtr hwndOwner, int nFolder, IntPtr hToken, int dwFlags, StringBuilder lpszPath);

    private const int MAX_PATH = 260;

    public enum eSHGFP_TYPE
    {
      SHGFP_TYPE_CURRENT  = 0,   // current value for user, verify it exists
      SHGFP_TYPE_DEFAULT  = 1,   // default value, may not exist
    };

    public enum eCSIDL
    {
      CSIDL_ADMINTOOLS                = 0x0030,        // <user name>\Start Menu\Programs\Administrative Tools
      CSIDL_APPDATA                   = 0x001a,        // <user name>\Application Data
      CSIDL_COMMON_ADMINTOOLS         = 0x002f,        // All Users\Start Menu\Programs\Administrative Tools
      CSIDL_COMMON_APPDATA            = 0x0023,        // All Users\Application Data
      CSIDL_COMMON_DOCUMENTS          = 0x002e,        // All Users\Documents
      CSIDL_COOKIES                   = 0x0021,
      CSIDL_FLAG_CREATE               = 0x8000,        // combine with CSIDL_ value to force folder creation in SHGetFolderPath()
      CSIDL_FLAG_DONT_VERIFY          = 0x4000,        // combine with CSIDL_ value to return an unverified folder path
      CSIDL_HISTORY                   = 0x0022,
      CSIDL_INTERNET_CACHE            = 0x0020,
      CSIDL_LOCAL_APPDATA             = 0x001c,        // <user name>\Local Settings\Applicaiton Data (non roaming)
      CSIDL_MYPICTURES                = 0x0027,        // C:\Program Files\My Pictures
      CSIDL_PERSONAL                  = 0x0005,        // My Documents
      CSIDL_PROGRAM_FILES             = 0x0026,        // C:\Program Files
      CSIDL_PROGRAM_FILES_COMMON      = 0x002b,        // C:\Program Files\Common
      CSIDL_SYSTEM                    = 0x0025,        // GetSystemDirectory()
      CSIDL_WINDOWS                   = 0x0024,        // GetWindowsDirectory()
    }

    #endregion //Types -----------------------------------------------------------------------------

    #region Fields
    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties
    #endregion //Properties ------------------------------------------------------------------------

    #region Private methods
    #endregion //Private methods -------------------------------------------------------------------

    #region Public methods

    public static FileSystemRights GetCurrentUsersFileSystemRights(string filePath)
    {
      WindowsIdentity user = WindowsIdentity.GetCurrent();
      IdentityReferenceCollection groups = user.Groups;
      AuthorizationRuleCollection aclRules = File.GetAccessControl(filePath)
                                                 .GetAccessRules(true, true, typeof(SecurityIdentifier));

      FileSystemRights allowedRights = (FileSystemRights)0,
                       deniedRights  = (FileSystemRights)0;

      foreach (FileSystemAccessRule rule in aclRules)
      {
        if (user.User == rule.IdentityReference)
        {
          #region user
          switch (rule.AccessControlType)
          {
            case AccessControlType.Allow:
              allowedRights |= rule.FileSystemRights;
              break;

            case AccessControlType.Deny:
              deniedRights |= rule.FileSystemRights;
              break;

            default:
              break;
          } //switch
          #endregion
        } //if

        foreach (IdentityReference group in groups)
        {
          if (group == rule.IdentityReference)
          {
            #region group
            switch (rule.AccessControlType)
            {
              case AccessControlType.Allow:
                allowedRights |= rule.FileSystemRights;
                break;

              case AccessControlType.Deny:
                deniedRights |= rule.FileSystemRights;
                break;

              default:
                break;
            } //switch
            #endregion
          } //if
        } //foreach
      } //foreach

      allowedRights &= ~deniedRights;

      return (allowedRights);
    }

    public static bool HasCurrentUserFileSystemRights(string filePath,
                                                      FileSystemRights rigths)
    {
      return ((GetCurrentUsersFileSystemRights(filePath) & rigths) == rigths);
    }


    public static string GetSpecialFolderPath(eSHGFP_TYPE flag,
                                              eCSIDL csidl)
    {
      StringBuilder sbPath = new StringBuilder(MAX_PATH);

      SHGetFolderPath(IntPtr.Zero, (int)csidl, IntPtr.Zero, (int)flag, sbPath);

      return (sbPath.ToString());
    }

    public static bool Contains(FileSystemRights rights,
                                FileSystemRights right)
    {
      return ((rights & right) == right);
    }

    public static bool Contains(FileAttributes rights,
                                FileAttributes right)
    {
      return ((rights & right) == right);
    }

    public static bool IsFileReadOnly(string file)
    {
      if (File.Exists(file))
        return (CUtil.Contains(File.GetAttributes(file), FileAttributes.ReadOnly));

      return (false);
    }

    #endregion //Public methods --------------------------------------------------------------------
  } //class
} //namespace
