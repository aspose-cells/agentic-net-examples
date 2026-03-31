using System;
using Aspose.Cells;

namespace WorkbookStructureProtectionDemo
{
    class Program
    {
        static void Main()
        {
            string inputPath = "OriginalWorkbook.xlsx";
            string protectedPath = "ProtectedWorkbook.xlsx";
            string unprotectedPath = "UnprotectedWorkbook.xlsx";
            string password = "MySecretPwd";

            // Create a new workbook and save it as the original file
            Workbook initialWb = new Workbook();
            initialWb.Worksheets[0].Name = "Sheet1";
            initialWb.Save(inputPath);
            initialWb.Dispose();

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the workbook structure
            workbook.Protect(ProtectionType.Structure, password);
            workbook.Save(protectedPath);
            workbook.Dispose();

            // Load the protected workbook with password
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook protectedWb = new Workbook(protectedPath, loadOptions);

            Console.WriteLine("Is workbook protected with password? " + protectedWb.IsWorkbookProtectedWithPassword);

            // Unprotect the workbook structure
            protectedWb.Unprotect(password);
            protectedWb.Save(unprotectedPath);
            protectedWb.Dispose();

            // Final verification
            Workbook finalWb = new Workbook(unprotectedPath);
            Console.WriteLine("Is workbook still protected? " + finalWb.IsWorkbookProtectedWithPassword);
            finalWb.Dispose();
        }
    }
}