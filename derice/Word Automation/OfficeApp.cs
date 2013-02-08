using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace derice.office
{
    public class OfficeApp
    {
        public string XML_Content { get; protected set; }

        public OfficeApp(string path)
        {
            XML_Content = System.IO.File.ReadAllText(path);
        }
        public OfficeApp(byte[] content)
        {
            XML_Content = Encoding.UTF8.GetString(content);
        }
    }
}
