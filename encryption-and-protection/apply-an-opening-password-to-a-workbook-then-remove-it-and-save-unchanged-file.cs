using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // 2. Apply an opening password to the workbook
            wb.Settings.Password = "OpenSecret";

            // 3. Save the password‑protected workbook
            string protectedPath = "protected_workbook.xlsx";
            wb.Save(protectedPath);
            Console.WriteLine($"Workbook saved with opening password: {protectedPath}");

            // 4. Load the protected workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "OpenSecret";
            Workbook protectedWb = new Workbook(protectedPath, loadOptions);
            Console.WriteLine("Protected workbook opened successfully.");

            // 5. Remove the opening password
            protectedWb.Settings.Password = null; // or string.Empty

            // 6. Save the workbook without a password (unchanged content)
            string unprotectedPath = "unprotected_workbook.xlsx";
            protectedWb.Save(unprotectedPath);
            Console.WriteLine($"Workbook saved without opening password: {unprotectedPath}");
        }
    }
}