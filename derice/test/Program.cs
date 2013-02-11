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
using System.Xml;
using System.Xml.Linq;

namespace test
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, string> elements = new Dictionary<string, string>();
            elements.Add("Value1", "new_value1");

            //Word app = new Word(@"C:\Users\Derice\Desktop\abc.xml");

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(@"C:\Users\Derice\Desktop\abc.xml");

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList nodeList = xmlDoc.SelectNodes("//w:tbl/w:tr/w:tc/w:p/w:r/w:t[text()='Value4']", namespaceManager);
            XmlNodeList nodeBookmark = xmlDoc.SelectNodes("//w:tbl/w:tr/w:tc/w:p/w:bookmarkStart[@w:name='tblTest']", namespaceManager);
            if (nodeBookmark.Count == 1) Console.WriteLine("Huat ar.."); Console.Read();
            if (nodeList.Count == 1)
            {
                XmlNode trNode = nodeList[0].ParentNode.ParentNode.ParentNode.ParentNode;
                XmlNode tblNode = trNode.ParentNode;
                XmlNode cloneNode = trNode.Clone();
                cloneNode.InnerXml = cloneNode.InnerXml.Replace("Value4", "new_Value4");
                tblNode.InsertBefore(cloneNode, tblNode.LastChild);
                xmlDoc.Save(@"C:\Users\Derice\Desktop\temp.xml");
            }
            
            Console.Write("done!");
            Console.Read();
        }

        protected XmlNode GetTrNodeInTableByKeyword(XmlDocument doc, string keyInTableCell)
        {
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList nodeList = doc.SelectNodes(string.Format("//w:tbl/w:tr/w:tc/w:p/w:r/w:t[text()='{0}']", keyInTableCell), namespaceManager);

            if (nodeList.Count > 0)
                return nodeList[0].ParentNode.ParentNode.ParentNode.ParentNode;
            else
                return null;
        }
    }
}
