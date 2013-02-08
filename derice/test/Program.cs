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
            int[] lucky_number = new int[3] { 1, 2, 3 };

            Person person1 = new Person();
            person1.IC = "86012814510322";
            person1.ContactNumber = 017663103422;
            person1.DOB = new DateTime(1986, 1, 28);
            person1.Name = "d3ric322";

            List<Person> ppl = new List<Person>();
            //ppl.Add(person);
            ppl.Add(person1);

            Person person = new Person();
            person.IC = "860128145103";
            person.ContactNumber = 0176631034;
            person.DOB = new DateTime(1986, 1, 28);
            person.Name = "d3ric3";
            person.LuckyNumber = ppl.ToArray();// lucky_number;
            person.Parents = ppl;
            

            //GetMyProperties(person);
            ExtractObjectProperties(person, "dd MMM yyyy");
            //PrintListObjectProperties(ConvertToObjectList(ppl));

            Console.Read();
        }

        protected static Dictionary<string, string> ExtractObjectProperties(Object obj, string DateTimeFormat = "dd/MM/yyyy")
        {
            Dictionary<string, string> rtnValues = new Dictionary<string, string>();

            foreach (PropertyInfo pi in obj.GetType().GetProperties())
            {
                if (pi.PropertyType.Name.Contains("List") || pi.PropertyType.Name.Contains(@"[]"))
                {
                    if(pi.GetValue(obj) != null)
                    foreach (var element in (IEnumerable)pi.GetValue(obj))
                    {
                        if (element.GetType().IsPrimitive || element.GetType() == typeof(decimal) || element.GetType() == typeof(string))
                            Console.WriteLine(string.Format("\t Name: {0}, Type: {1}, Value: {2}", pi.Name, element.GetType().Name, element.ToString()));
                        else
                            ExtractObjectPrimitiveProperties(element, DateTimeFormat);
                    }
                }                
                else
                {
                    if (pi.PropertyType.Name == "DateTime")
                    {
                        string key = pi.Name;
                        string value = ((DateTime)pi.GetValue(obj, null)).ToString(DateTimeFormat);
                        rtnValues.Add(key, value);
                        Console.WriteLine(string.Format("Name: {0}, Type: {1}, Value: {2}", key, pi.PropertyType.Name, value));
                    }
                    else
                    {
                        string key = pi.Name;
                        string value = pi.GetValue(obj, null).ToString();
                        rtnValues.Add(key, value);
                        Console.WriteLine(string.Format("Name: {0}, Type: {1}, Value: {2}", key, pi.PropertyType.Name, value));
                    }
                }                
            }

            return rtnValues;
        }

        protected static Dictionary<string, string> ExtractObjectPrimitiveProperties(Object obj, string DateTimeFormat = "dd/MM/yyyy")
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
                        Console.WriteLine(string.Format("\t Name: {0}, Type: {1}, Value: {2}", key, pi.PropertyType.Name, value));
                    }
                    else
                    {
                        string key = pi.Name;
                        string value = pi.GetValue(obj, null).ToString();
                        rtnValues.Add(key, value);
                        Console.WriteLine(string.Format("\t Name: {0}, Type: {1}, Value: {2}", key, pi.PropertyType.Name, value));
                    }
                }
            }

            return rtnValues;
        }

        protected static List<Dictionary<string, string>> ExtractListObjectProperties(List<object> objects, string DateTimeFormat = "dd/MM/yyyy")
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

        public static void GetMyProperties(object obj)
        {
            foreach (PropertyInfo pinfo in obj.GetType().GetProperties())
            {
                var getMethod = pinfo.GetGetMethod();
                if (getMethod.ReturnType.IsArray)
                {
                    var arrayObject = getMethod.Invoke(obj, null);
                    foreach (object element in (Array)arrayObject)
                    {
                        Console.WriteLine(element.ToString());
                        //foreach (PropertyInfo arrayObjPinfo in element.GetType().GetProperties())
                        //{
                        //    Console.WriteLine(arrayObjPinfo.Name + ":" + arrayObjPinfo.GetGetMethod().Invoke(element, null).ToString());
                        //}
                    }
                }
            }
        }
    }


    public class Person
    {
        public string IC { get; set; }
        public Int64 ContactNumber { get; set; }
        public DateTime DOB { get; set; }
        public string Name { get; set; }
        public List<Person> Parents { get; set; }
        public Person[] LuckyNumber { get; set; }
    }
}
