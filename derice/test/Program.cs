using derice.office;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace test
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, string> elements = new Dictionary<string, string>();
            elements.Add("Value1", "new_value1");

            Word app = new Word(@"C:\Users\Derice\Desktop\abc.xml");
            string content = app.ReplaceKeywords(elements, null);
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter(@"C:\Users\Derice\Desktop\temp.xml"))
            {
                sw.Write(content);
            }
            Console.Write("done!");
            Console.Read();
        }
    }
}
