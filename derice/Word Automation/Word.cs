using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace derice.office
{
    public class Word : OfficeApp
    {
        public Word(string path) : base(path) { }
        public Word(byte[] content) : base(content) { }

        public string ReplaceKeywords(Dictionary<string, string> dicKeyValuePairs, DataSet dsTabularValues = null)
        {
            StringBuilder sb = new StringBuilder(XML_Content);

            //replace flat data
            foreach (var element in dicKeyValuePairs)
            {
                sb = sb.Replace(string.Format("[{0}]", element.Key), element.Value);
            }

            //replace tabular data in a table
            if (dsTabularValues != null)
            {
                for (int i = 0; i < dsTabularValues.Tables.Count; i++)
                {
                    sb = ReplaceTabularKeywords(sb, dsTabularValues.Tables[i]);
                }
            }

            return sb.ToString();
        }

        protected StringBuilder ReplaceTabularKeywords(StringBuilder sbContent, DataTable dtKeyValue)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(sbContent.ToString());

            //get table row (w:tr) by bookmark within table cell
            XmlNode trNode = null;
            trNode = GetTrNodeInTableByBookmarkName(doc, dtKeyValue.TableName);

            //if table row (w:tr) not found, continue to get table row by keywords
            if (trNode == null)
            {
                foreach (DataColumn column in dtKeyValue.Columns)
                {
                    trNode = GetTrNodeInTableByKeyword(doc, string.Format("[{0}]", column.ColumnName));

                    if (trNode != null) break;
                }
            }

            //if table row found
            if (trNode != null)
            {                
                XmlNode tblNode = trNode.ParentNode;

                foreach (DataRow dr in dtKeyValue.Rows)
                {
                    XmlNode cloneTrNode = trNode.Clone();
                    foreach (DataColumn column in dtKeyValue.Columns)
                    {
                        string key = String.Format("[{0}]", column.ColumnName);
                        string value = dr[column.ColumnName].ToString();
                        cloneTrNode.InnerXml = cloneTrNode.InnerXml.Replace(key, value);                        
                    }
                    tblNode.InsertBefore(cloneTrNode, tblNode.LastChild);
                }

                tblNode.RemoveChild(trNode);
            }

            //convert modified XmlDocument to string builder
            using (StringWriter sw = new StringWriter())
            {
                using (XmlWriter xmltw = XmlWriter.Create(sw))
                {
                    doc.WriteTo(xmltw);
                    xmltw.Flush();
                    sbContent = sw.GetStringBuilder();                  
                }
            }

            return sbContent;
        }

        //get table row (w:tr) by keyword provided
        protected XmlNode GetTrNodeInTableByKeyword(XmlDocument doc, string keywordInTableCell)
        {            
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList nodeList = doc.SelectNodes(string.Format("//w:tbl/w:tr/w:tc/w:p/w:r/w:t[text()='{0}']", keywordInTableCell), namespaceManager);

            if (nodeList.Count > 0)
                return nodeList[0].ParentNode.ParentNode.ParentNode.ParentNode;
            else
                return null;
        }

        //get table row (w:tr) by bookmark that is within same cell with keyword
        protected XmlNode GetTrNodeInTableByBookmarkName(XmlDocument doc, string bookmarkName)
        {
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList nodeList = doc.SelectNodes(string.Format("//w:tbl/w:tr/w:tc/w:p/w:bookmarkStart[@w:name='{0}']", bookmarkName), namespaceManager);

            if (nodeList.Count > 0)
                return nodeList[0].ParentNode.ParentNode.ParentNode;
            else
                return null;
        }

        #region backup
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

        //extract one object's properties 
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

        protected static List<Object> ConvertToObjectList(IEnumerable<Object> objs)
        {
            List<object> value = new List<object>();
            foreach (var obj in objs)
            {
                value.Add(obj);
            }
            return value;
        }
        #endregion
    }
}
