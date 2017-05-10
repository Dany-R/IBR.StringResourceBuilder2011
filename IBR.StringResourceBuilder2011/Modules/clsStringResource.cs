using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;



namespace IBR.StringResourceBuilder2011.Modules
{
  /// <summary>
  /// StringResource class which holds a resource name, string literal and its location.
  /// </summary>
  class StringResource
  {
    #region Constructor

    /// <summary>
    /// Initializes a new instance of the <see cref="StringResource"/> class.
    /// </summary>
    public StringResource()
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StringResource"/> class.
    /// </summary>
    /// <param name="name">The resource name.</param>
    /// <param name="text">The string literal.</param>
    /// <param name="location">The location.</param>
    public StringResource(string name,
                          string text,
                          System.Drawing.Point location)
    {
      this.Name     = name;
      this.Text     = text;
      this.Location = location;
    }

    #endregion //Constructor -----------------------------------------------------------------------

    #region Types
    #endregion //Types -----------------------------------------------------------------------------

    #region Fields
    #endregion //Fields ----------------------------------------------------------------------------

    #region Properties

    /// <summary>
    /// Gets or sets the resource name.
    /// </summary>
    /// <value>
    /// The resource name.
    /// </value>
    public string Name { get; set; }

    /// <summary>
    /// Gets the string literal (text).
    /// </summary>
    public string Text { get; private set; }

    //must be defined explicitly because otherwise Offset() would not work!
    private System.Drawing.Point m_Location = System.Drawing.Point.Empty;
    /// <summary>
    /// Gets the location of the string literal (X is line number and Y is column number).
    /// </summary>
    public System.Drawing.Point Location { get { return (m_Location); } private set { m_Location = value; } }

    #endregion //Properties ------------------------------------------------------------------------

    #region Events
    #endregion //Events ----------------------------------------------------------------------------

    #region Private methods
    #endregion //Private methods -------------------------------------------------------------------

    #region Public methods

    /// <summary>
    /// Translates the location by the specified amount.
    /// </summary>
    /// <param name="lineOffset">The line offset.</param>
    /// <param name="columnOffset">The column offset.</param>
    public void Offset(int lineOffset,
                       int columnOffset)
    {
      if ((lineOffset == 0) && (columnOffset == 0))
        return;

      m_Location.Offset(lineOffset, columnOffset);
    }

    #endregion //Public methods --------------------------------------------------------------------
  } //class
} //namespace
