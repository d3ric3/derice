using derice.office;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
            Word word = new Word(@"C:\Users\Derice\Desktop\abc.xml");
            Dictionary<string, string> user = new Dictionary<string, string>();
            user.Add("name", "Derice Kong");
            user.Add("gender", "Male");
            user.Add("email", "d3ric3@gmail.com");

            DataTable dt = new DataTable();
            dt.Columns.Add("product");
            dt.Columns.Add("price");

            dt.TableName = "tblTest";

            DataRow dr1 = dt.NewRow();
            dr1[0] = "i7 3770";
            dr1[1] = "1200";
            dt.Rows.Add(dr1);

            DataRow dr2 = dt.NewRow();
            dr2[0] = "i5 3310";
            dr2[1] = "900";
            dt.Rows.Add(dr2);

            DataRow dr3 = dt.NewRow();
            dr3[0] = "i3 3310";
            dr3[1] = "500";
            dt.Rows.Add(dr3);

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);

            string xmlContent = word.ReplaceKeywords(user, ds);

            if (!File.Exists(@"C:\Users\Derice\Desktop\result.xml"))
            {
                using(File.Create(@"C:\Users\Derice\Desktop\result.xml")){}
            }

            using (TextWriter tw = new StreamWriter(@"C:\Users\Derice\Desktop\result.xml"))
            {
                tw.Write(xmlContent);
            }
            
            Console.Write("done!");
            Console.Read();
        }

        protected void Backup()
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
