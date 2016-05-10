using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = @".\test.xls";
            var excelReader = new ExcelReader(fileName);
            excelReader.WriteContent();

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
