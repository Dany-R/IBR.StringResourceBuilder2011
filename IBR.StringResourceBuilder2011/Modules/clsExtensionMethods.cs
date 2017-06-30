using System;
using System.Windows;
using System.Windows.Threading;
using EnvDTE;

namespace IBR.StringResourceBuilder2011.Modules
{
  public static class ExtensionMethods
  {
    #region UIElement.Refresh

    private static Action EmptyDelegate = delegate() { };

    public static void Refresh(this UIElement uiElement)
    {
      uiElement.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);
    }

    #endregion

    #region CodeElement

    public static bool HasStartPoint(this CodeElement element)
    {
      try
      {
        //test whether access to the StartPoint throws an exception
        int line = element.StartPoint.Line;

        return (true);
      }
      catch (Exception /*ex*/)
      {
        return (false);
      }
    }

      #endregion
      //public static class TextBoxWatermarkExtensionMethod
      //{
      //  private const uint ECM_FIRST = 0x1500;
      //  private const uint EM_SETCUEBANNER = ECM_FIRST + 1;

      //  [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
      //  private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

      //  public static void SetWatermark(this TextBox textBox, string watermarkText)
      //  {
      //    SendMessage(textBox.Handle, EM_SETCUEBANNER, 0, watermarkText);
      //  }
      //}
    } //class
} //namespace
