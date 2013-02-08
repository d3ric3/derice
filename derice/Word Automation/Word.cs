using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace derice.office
{
    public class Word : OfficeApp
    {
        public Word(string path) : base(path) { }
        public Word(byte[] content) : base(content) { }

        protected Dictionary<string,string> ExtractMacroButton()
        {
            Dictionary<string, string> rtnValues = new Dictionary<string, string>();

            //Retrieve all occurance in the format of {MACROBUTTON NoMacro [keyword]}
            foreach (Match m in Regex.Matches(this.XML_Content, @"\{\s{0,}macrobutton\s{1,}nomacro\s{1,}\[[_\w\d]+\]\s{0,}\}", RegexOptions.IgnoreCase))
            {                
                //retrieve value between [ to ]
                int startIndexOfKey = m.ToString().IndexOf('[') + 1;
                int lengthOfKey = m.ToString().IndexOf(']') - m.ToString().IndexOf('[') - 1;

                string key = m.ToString().Substring(startIndexOfKey, lengthOfKey);
                string value = m.ToString();
                rtnValues.Add(key, value);
            }

            return rtnValues;
        }

        protected Dictionary<string,string> ExtractObjectProperties(Object obj, string DateTimeFormat = "dd/MM/yyyy")
        {
            Dictionary<string, string> rtnValues = new Dictionary<string, string>();
            
            foreach (PropertyInfo pi in obj.GetType().GetProperties())
            {
                if (pi.PropertyType.Name.Contains("List") || pi.PropertyType.Name.Contains(@"[]"))
                {
                    continue;
                }
                else
                {
                    if (pi.PropertyType.Name == "DateTime")
                    {
                        string key = pi.Name;
                        string value = ((DateTime)pi.GetValue(obj, null)).ToString(DateTimeFormat);
                        rtnValues.Add(key, value);                        
                    }
                    else
                    {
                        string key = pi.Name;
                        string value = pi.GetValue(obj, null).ToString();
                        rtnValues.Add(key, value);
                    }
                }
                //Console.WriteLine(string.Format("Name: {0}, Type: {1}, Value: {2}", pi.Name, pi.PropertyType.Name, pi.GetValue(obj, null)));
            }

            return rtnValues;
        }

        protected List<Dictionary<string, string>> ExtractListObjectProperties(List<object> objects, string DateTimeFormat = "dd/MM/yyyy")
        {
            List<Dictionary<string, string>> rtnValues = new List<Dictionary<string, string>>();

            foreach (var obj in objects)
            {
                rtnValues.Add(ExtractObjectProperties(obj, DateTimeFormat));
            }

            return rtnValues;
        }
    }
}
