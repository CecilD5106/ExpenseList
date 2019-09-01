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
        public class Expenses
        {
            //Constructor
            public Expenses()
            {
                Alcohol = 0.00;
                Cash = 0.00;
                Clothing = 0.00;
                ClothingCare = 0.00;
                ComputerBusinessExpense = 0.00;
                DuesOrganization = 0.00;
                DuesProperty = 0.00;
                Education = 0.00;
                EducationCollege = 0.00;
                Electronics = 0.00;
                Entertainment = 0.00;
                FeesBanking = 0.00;
                FoodDining = 0.00;
                FoodFastFood = 0.00;
                FoodGroceries = 0.00;
                FoodLunch = 0.00;
                FoodSnack = 0.00;
                Gas = 0.00;
                Gifts = 0.00;
                InsuranceAuto = 0.00;
                InsuranceLife = 0.00;
                InterestCharged = 0.00;
                LicensingAuto = 0.00;
                MaintenanceAuto = 0.00;
                MaintenanceHome = 0.00;
                Medical = 0.00;
                Merchandise = 0.00;
                Mortgage = 0.00;
                Other = 0.00;
                Parking = 0.00;
                PaymentAuto = 0.00;
                PaymentCreditCard = 0.00;
                Postage = 0.00;
                ServiceFee = 0.00;
                Subscription = 0.00;
                SubscriptionComputer = 0.00;
                SubscriptionFitness = 0.00;
                TaxesFederalIncome = 0.00;
                TaxesProperty = 0.00;
                TaxesSales = 0.00;
                TaxesStateIncome = 0.00;
                Travel = 0.00;
                UtilityElectric = 0.00;
                UtilityGas = 0.00;
                UtilityPhone = 0.00;
                UtilitySatellite = 0.00;
                UtilityWater = 0.00;
            }

            //Properties
            public double Alcohol { get; private set; }
            public double Cash { get; private set; }
            public double Clothing { get; private set; }
            public double ClothingCare { get; private set; }
            public double ComputerBusinessExpense { get; private set; }
            public double DuesOrganization { get; private set; }
            public double DuesProperty { get; private set; }
            public double Education { get; private set; }
            public double EducationCollege { get; private set; }
            public double Electronics { get; private set; }
            public double Entertainment { get; private set; }
            public double FeesBanking { get; private set; }
            public double FoodDining { get; private set; }
            public double FoodFastFood { get; private set; }
            public double FoodGroceries { get; private set; }
            public double FoodLunch { get; private set; }
            public double FoodSnack { get; private set; }
            public double Gas { get; private set; }
            public double Gifts { get; private set; }
            public double InsuranceAuto { get; private set; }
            public double InsuranceLife { get; private set; }
            public double InterestCharged { get; private set; }
            public double LicensingAuto { get; private set; }
            public double MaintenanceAuto { get; private set; }
            public double MaintenanceHome { get; private set; }
            public double Medical { get; private set; }
            public double Merchandise { get; private set; }
            public double Mortgage { get; private set; }
            public double Other { get; private set; }
            public double Parking { get; private set; }
            public double PaymentAuto { get; private set; }
            public double PaymentCreditCard { get; private set; }
            public double Postage { get; private set; }
            public double ServiceFee { get; private set; }
            public double Subscription { get; private set; }
            public double SubscriptionComputer { get; private set; }
            public double SubscriptionFitness { get; private set; }
            public double TaxesFederalIncome { get; private set; }
            public double TaxesProperty { get; private set; }
            public double TaxesSales { get; private set; }
            public double TaxesStateIncome { get; private set; }
            public double Travel { get; private set; }
            public double UtilityElectric { get; private set; }
            public double UtilityGas { get; private set; }
            public double UtilityPhone { get; private set; }
            public double UtilitySatellite { get; private set; }
            public double UtilityWater { get; private set; }

            public void AddExpenses(string sExpense, double dExpense)
            {
                switch (sExpense)
                {
                    case "Alcohol":
                        Alcohol += dExpense;
                        break;
                    case "Cash":
                        Cash += dExpense;
                        break;
                    case "Clothing":
                        Clothing += dExpense;
                        break;
                    case "Clothing - Care":
                        ClothingCare += dExpense;
                        break;
                    case "Computer Business - Expense":
                        ComputerBusinessExpense += dExpense;
                        break;
                    case "Dues - Organization":
                        DuesOrganization += dExpense;
                        break;
                    case "Dues - Property":
                        DuesProperty += dExpense;
                        break;
                    case "Education":
                        Education += dExpense;
                        break;
                    case "Education - College":
                        EducationCollege += dExpense;
                        break;
                    case "Electronics":
                        Electronics += dExpense;
                        break;
                    case "Entertainment":
                        Entertainment += dExpense;
                        break;
                    case "Fees - Banking":
                        FeesBanking += dExpense;
                        break;
                    case "Food - Dining":
                        FoodDining += dExpense;
                        break;
                    case "Food - Fast Food":
                        FoodFastFood += dExpense;
                        break;
                    case "Food - Groceries":
                        FoodGroceries += dExpense;
                        break;
                    case "Food - Lunch":
                        FoodLunch += dExpense;
                        break;
                    case "Food - Snacks":
                        FoodSnack += dExpense;
                        break;
                    case "Gas":
                        Gas += dExpense;
                        break;
                    case "Gifts":
                        Gifts += dExpense;
                        break;
                    case "Insurance - Auto":
                        InsuranceAuto += dExpense;
                        break;
                    case "Insurance - Life":
                        InsuranceLife += dExpense;
                        break;
                    case "Interest - Charged":
                        InterestCharged += dExpense;
                        break;
                    case "Licensing - Auto":
                        LicensingAuto += dExpense;
                        break;
                    case "Maintenance - Auto":
                        MaintenanceAuto += dExpense;
                        break;
                    case "Maintenance - Home":
                        MaintenanceHome += dExpense;
                        break;
                    case "Medical":
                        Medical += dExpense;
                        break;
                    case "Merchandise":
                        Merchandise += dExpense;
                        break;
                    case "Mortgage":
                        Mortgage += dExpense;
                        break;
                    case "Other":
                        Other += dExpense;
                        break;
                    case "Parking":
                        Parking += dExpense;
                        break;
                    case "Payment - Auto":
                        PaymentAuto += dExpense;
                        break;
                    case "Payment - Credit Card":
                        PaymentCreditCard += dExpense;
                        break;
                    case "Postage":
                        Postage += dExpense;
                        break;
                    case "Service Fee":
                        ServiceFee += dExpense;
                        break;
                    case "Subscription":
                        Subscription += dExpense;
                        break;
                    case "Subscription - Computer":
                        SubscriptionComputer += dExpense;
                        break;
                    case "Subscription - Fitness":
                        SubscriptionFitness += dExpense;
                        break;
                    case "Taxes - Federal Income":
                        TaxesFederalIncome += dExpense;
                        break;
                    case "Taxes - Property":
                        TaxesProperty += dExpense;
                        break;
                    case "Taxes - Sales":
                        TaxesSales += dExpense;
                        break;
                    case "Taxes - State Income":
                        TaxesStateIncome += dExpense;
                        break;
                    case "Travel":
                        Travel += dExpense;
                        break;
                    case "Utility - Electric":
                        UtilityElectric += dExpense;
                        break;
                    case "Utility - Gas":
                        UtilityGas += dExpense;
                        break;
                    case "Utility - Phone":
                        UtilityPhone += dExpense;
                        break;
                    case "Utility - Satellite":
                        UtilitySatellite += dExpense;
                        break;
                    case "Utility - Water":
                        UtilityWater += dExpense;
                        break;
                }
            }
        }
        static void Main(string[] args)
        {
            //Set Variables
            string path = "E:\\Excel\\Bank Accounts - Copy.xlsx";
            List<string> sTabs = new List<string>(new string[] { "Major Bills", "Family", "Savings", "Stock", "Visa" });
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(path);
            try
            {
                Console.WriteLine("Enter the start date.\n");
                DateTime dtStart = Convert.ToDateTime(Console.ReadLine());
                Console.WriteLine("Enter the end date.\n");
                DateTime dtEnd = Convert.ToDateTime(Console.ReadLine());
                Expenses monthlyExpenses = new Expenses();
                foreach (string sTab in sTabs)
                {
                    Worksheet ws = wb.Worksheets[sTab];
                    int maxRow = 0;
                    int i = 3;
                    while (ws.Cells[ i, 1].Value != null)
                    {
                        DateTime dtExpense = Convert.ToDateTime(ws.Cells[i, 1].Value);
                        if (dtStart <= dtExpense && dtExpense <= dtEnd)
                        {
                            string sExpense = ws.Cells[i, 10].Value;
                            double dExpense = 0.00;
                            if (ws.Cells[i, 5].Value != null)
                            {
                                dExpense = Convert.ToDouble(ws.Cells[i, 5].Value);
                            }
                            monthlyExpenses.AddExpenses(sExpense, dExpense);
                        }
                        i++;
                    }
                    maxRow = i;
                }
                Worksheet wsExp = wb.Worksheets["Expenses"];
                int j = 1;
                while (wsExp.Cells[2, j].Value != null)
                {
                    j++;
                }
                for (int k = 2; k < 50; k++)
                {
                    switch (k)
                    {
                        case 2:
                            wsExp.Cells[k, j].Value = "Total";
                            break;
                        case 3:
                            wsExp.Cells[k, j].Value = monthlyExpenses.Alcohol;
                            break;
                        case 4:
                            wsExp.Cells[k, j].Value = monthlyExpenses.Cash;
                            break;
                        case 5:
                            wsExp.Cells[k, j].Value = monthlyExpenses.Clothing;
                            break;
                        case 6:
                            wsExp.Cells[k, j].Value = monthlyExpenses.ClothingCare;
                            break;
                        case 7:
                            wsExp.Cells[k, j].Value = monthlyExpenses.ComputerBusinessExpense;
                            break;
                        case 8:
                            wsExp.Cells[k, j].Value = monthlyExpenses.DuesOrganization;
                            break;
                        case 9:
                            wsExp.Cells[k, j].Value = monthlyExpenses.DuesProperty;
                            break;
                        case 10:
                            wsExp.Cells[k, j].Value = monthlyExpenses.Education;
                            break;
                    }
                }
                wb.Save();
                excel.Quit();
            }
            catch (Exception e)
            {
                excel.Quit();
                Console.WriteLine(e.ToString());
                throw;
            } finally
            {
                excel.Quit();
            }
        }
    }
}
