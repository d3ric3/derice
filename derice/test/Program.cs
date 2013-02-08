using derice.office;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace test
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime dt = new DateTime(2013, 2, 1);
            dt = dt.AddDays(60);

            Console.WriteLine(DateTime.Today.AddDays(-30).ToString("dd-MMM-yyyy"));
            Console.Read();
        }
    }
}
