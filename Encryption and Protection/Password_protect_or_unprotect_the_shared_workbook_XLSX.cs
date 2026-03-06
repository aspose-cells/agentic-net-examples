using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SharedWorkbookProtectionDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Protect the shared workbook with a password
            wb.ProtectSharedWorkbook("myPassword");

            // Save the protected workbook
            wb.Save("ProtectedSharedWorkbook.xlsx");

            // Load the protected workbook
            Workbook wbProtected = new Workbook("ProtectedSharedWorkbook.xlsx");

            // Verify protection status
            Console.WriteLine("Workbook is protected: " + wbProtected.Settings.IsProtected);

            // Unprotect the shared workbook with the same password
            wbProtected.UnprotectSharedWorkbook("myPassword");

            // Save the unprotected workbook
            wbProtected.Save("UnprotectedSharedWorkbook.xlsx");

            // Verify that the workbook is no longer protected
            Workbook wbUnprotected = new Workbook("UnprotectedSharedWorkbook.xlsx");
            Console.WriteLine("Workbook is protected after unprotect: " + wbUnprotected.Settings.IsProtected);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SharedWorkbookProtectionDemo.Run();
        }
    }
}