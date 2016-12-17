using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ReadExcelFile
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// 
        /// need to make changes
        /// </summary>
        [STAThread]
        static void Main()
        {
            

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range2;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Test\test.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;

            List<employeeFile1> allEmp1 = new List<employeeFile1>();



            for (int i = 2; i <= rowCount; i++)
            {
                

                System.Array MyValues = (System.Array)xlWorkSheet.get_Range("A" + i.ToString(), "Q" + i.ToString()).Cells.Value;

                //if (range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null)
                employeeFile1 newEmp = new employeeFile1();

                newEmp.employeeSIN = MyValues.GetValue(1, 1) == null ? "" : MyValues.GetValue(1, 1).ToString();
                newEmp.employeeLastName = MyValues.GetValue(1, 2) == null ? "" : MyValues.GetValue(1, 2).ToString();
                newEmp.employeeFirstName = MyValues.GetValue(1, 3) == null ? "" : MyValues.GetValue(1, 3).ToString();
                newEmp.employeeMiddleName = MyValues.GetValue(1, 4) == null ? "" : MyValues.GetValue(1, 4).ToString();
                newEmp.employeeBirthday = MyValues.GetValue(1, 5) == null ? new DateTime(1900,01,01) : (DateTime)MyValues.GetValue(1, 5);
                newEmp.employeeEIOccupation = MyValues.GetValue(1, 6) == null ? "" : MyValues.GetValue(1, 6).ToString();
                newEmp.companyBranch = MyValues.GetValue(1, 7) == null ? "" : MyValues.GetValue(1, 7).ToString();
                newEmp.companyBrachName = MyValues.GetValue(1, 8) == null ? "" : MyValues.GetValue(1, 8).ToString();
                newEmp.companyDept = MyValues.GetValue(1, 9) == null ? "" : MyValues.GetValue(1, 9).ToString();
                newEmp.companyDeptName = MyValues.GetValue(1, 10) == null ? "" : MyValues.GetValue(1, 10).ToString();
                newEmp.employeeHireDate = MyValues.GetValue(1, 11) == null ? new DateTime(1900, 01, 01) : (DateTime)MyValues.GetValue(1, 11);
                newEmp.companyCompany = MyValues.GetValue(1, 12) == null ? "" : MyValues.GetValue(1, 12).ToString();
                newEmp.employeePayrollStatus = MyValues.GetValue(1, 13) == null ? "" : MyValues.GetValue(1, 13).ToString();
                newEmp.employeePaymentType = MyValues.GetValue(1, 14) == null ? "" : MyValues.GetValue(1, 14).ToString();
                newEmp.employeeRate = MyValues.GetValue(1, 15) == null ? 0 : Convert.ToDouble(MyValues.GetValue(1, 15).ToString());
                newEmp.employeeSalaryPerPay = MyValues.GetValue(1, 16) == null ? 0 : (double)MyValues.GetValue(1, 16);
                newEmp.employeeStandardHours = MyValues.GetValue(1, 17) == null ? 0 : (double)MyValues.GetValue(1, 17);



                allEmp1.Add(newEmp);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());


            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(xlWorkSheet);

            //close and release
            xlWorkBook.Close();
            Marshal.ReleaseComObject(xlWorkBook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
