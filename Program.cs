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

            try
            {
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
