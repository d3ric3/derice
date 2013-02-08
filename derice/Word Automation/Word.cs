using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace derice.office
{
    public class Word : OfficeApp
    {
        public Word(string path) : base(path) { }
        public Word(byte[] content) : base(content) { }

        protected string[] ExtractMacroButton()
        {
            List<string> keys = new List<string>();

            //Retrieve all occurance in the format of {MACROBUTTON NoMacro [keyword]}
            foreach (Match m in Regex.Matches(this.XML_Content, @"\{\s{0,}macrobutton\s{1,}nomacro\s{1,}\[[_\w\d]+\]\s{0,}\}", RegexOptions.IgnoreCase))
            {                
                //retrieve value between [ to ]
                int startIndexOfKey = m.ToString().IndexOf('[') + 1;
                int lengthOfKey = m.ToString().IndexOf(']') - m.ToString().IndexOf('[') - 1;
                string key = m.ToString().Substring(startIndexOfKey, lengthOfKey);
            }

            return keys.ToArray();
        }
    }
}
