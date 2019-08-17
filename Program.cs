using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExpenseList
{
    class Program
    {
        static void Main(string[] args)
        {
            //Set Variables
            string path = "E:\\Excel\\Bank Accounts.xlsx";
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(path);
            List<string> sTabs = new List<string>(new string[] { "Major Bills", "Family", "Savings", "Stock", "Visa" });
            try
            {
                foreach (string sTab in sTabs)
                {
                    Worksheet ws = wb.Worksheets[sTab];
                    Worksheet wsExp = wb.Worksheets["Expenses"];
                }
                wb.Save();
                excel.Quit();
            }
            catch (Exception)
            {
                excel.Quit();
                throw;
            } finally
            {
                excel.Quit();
            }
        }
    }
}
