using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectSharedWorkbookWriteProtectionDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data
            wb.Worksheets[0].Cells["A1"].PutValue("Shared and Write Protected");

            // Enable sharing for the workbook
            wb.Settings.Shared = true;

            // Apply write protection to the shared workbook with a password
            wb.ProtectSharedWorkbook("myPassword");

            // Save the protected, shared workbook
            string outputPath = "SharedWriteProtectedWorkbook.xlsx";
            wb.Save(outputPath);

            // Load the saved workbook to verify protection and sharing settings
            Workbook loadedWb = new Workbook(outputPath);

            // Verify that sharing is still enabled
            Console.WriteLine("Is workbook shared? " + loadedWb.Settings.Shared);

            // Verify that the shared workbook is protected
            Console.WriteLine("Is workbook protected? " + loadedWb.Settings.IsProtected);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectSharedWorkbookWriteProtectionDemo.Run();
        }
    }
}