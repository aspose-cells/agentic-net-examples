using System;
using Aspose.Cells;

namespace AsposeCellsPasswordProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Apply password protection (encrypts the workbook)
            wb.Settings.Password = "StrongPassword123";

            // Save the protected workbook
            string protectedPath = "ProtectedWorkbook.xlsx";
            wb.Save(protectedPath, SaveFormat.Xlsx);

            // Load the password‑protected workbook using LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "StrongPassword123";
            Workbook wbLoaded = new Workbook(protectedPath, loadOptions);

            // Verify that the data can be accessed after providing the password
            string cellValue = wbLoaded.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Loaded cell value: " + cellValue);
        }
    }
}