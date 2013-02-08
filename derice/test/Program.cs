using derice.office;
using System;
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
            Person person = new Person();
            person.IC = "860128145103";
            person.ContactNumber = 0176631034;
            person.DOB = new DateTime(1986, 1, 28);
            person.Name = "d3ric3";

            Person person1 = new Person();
            person1.IC = "86012814510322";
            person1.ContactNumber = 017663103422;
            person1.DOB = new DateTime(1986, 1, 28);
            person1.Name = "d3ric322";

            List<Person> ppl = new List<Person>();
            ppl.Add(person);
            ppl.Add(person1);

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
                    continue;
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

        protected static void PrintListObjectProperties(List<Object> objs)
        {
            foreach (var obj in objs)
            {
                ExtractObjectProperties(obj);
            }
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
    }


    public class Person
    {
        public string IC { get; set; }
        public Int64 ContactNumber { get; set; }
        public DateTime DOB { get; set; }
        public string Name { get; set; }
        public List<Person> Buddies { get; set; }
        public Person[] MyProperty { get; set; }
    }
}
