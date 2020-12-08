using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string[][] updateData = { new string[] { "1", "CGTN", "频道xxx","纪实", "197", "5.0", "8.2%" }, new string[] { "2", "CGTN", "频道", "纪实", "197", "5.0", "0.2%" } };
            string path = Environment.CurrentDirectory + "\\Template.xlsx";
            NPOIHelper.UpdateExcel2(path, "sheet1", updateData, 1);
        }
    }
}
