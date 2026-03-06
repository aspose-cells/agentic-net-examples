using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectUnprotectSharedWorkbookDemo
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook wb = new Workbook();

            // Protect the shared workbook with a password
            wb.ProtectSharedWorkbook("myPassword");

            // Save the protected workbook (lifecycle: save)
            wb.Save("ProtectedSharedWorkbook.xlsx");

            // Load the protected workbook (lifecycle: load)
            Workbook wbProtected = new Workbook("ProtectedSharedWorkbook.xlsx");

            // Verify that the workbook is protected
            Console.WriteLine("Workbook is protected: " + wbProtected.Settings.IsProtected);

            // Unprotect the shared workbook using the same password
            wbProtected.UnprotectSharedWorkbook("myPassword");

            // Save the unprotected workbook
            wbProtected.Save("UnprotectedSharedWorkbook.xlsx");

            // Clean up
            wb.Dispose();
            wbProtected.Dispose();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ProtectUnprotectSharedWorkbookDemo.Run();
        }
    }
}