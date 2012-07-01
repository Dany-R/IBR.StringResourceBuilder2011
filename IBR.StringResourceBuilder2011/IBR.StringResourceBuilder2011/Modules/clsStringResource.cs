using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;



namespace IBR.StringResourceBuilder2011.Modules
{
  class StringResource
  {
    public StringResource()
    {
    }

    public StringResource(string name,
                          string text,
                          System.Drawing.Point location)
    {
      this.Name = name;
      this.Text = text;
      this.Location = location;
    }

    //must be defined explicitly because otherwise Offset() would not work!
    private System.Drawing.Point m_Location = System.Drawing.Point.Empty;

    public string Name { get; set; }
    public string Text { get; private set; }
    public System.Drawing.Point Location { get { return (m_Location); } private set { m_Location = value; } }

    public void Offset(int dx, int dy)
    {
      m_Location.Offset(dx, dy);
    }
  }
}
