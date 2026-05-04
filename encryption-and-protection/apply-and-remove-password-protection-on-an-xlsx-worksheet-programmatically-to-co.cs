using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(12345);

            // Apply password protection to the worksheet
            // Protect all possible actions with the password "mySecret"
            sheet.Protect(ProtectionType.All, "mySecret", null);

            // Save the protected workbook
            workbook.Save("ProtectedWorksheet.xlsx");

            // Load the protected workbook
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = null; // No workbook-level password, only sheet protection
            Workbook protectedWb = new Workbook("ProtectedWorksheet.xlsx", loadOptions);
            Worksheet protectedSheet = protectedWb.Worksheets[0];

            // Verify that the sheet is protected
            Console.WriteLine("Worksheet is protected: " + protectedSheet.IsProtected);

            // Remove the password protection using the correct password
            protectedSheet.Unprotect("mySecret");

            // Verify that the sheet is now unprotected
            Console.WriteLine("Worksheet is protected after unprotect: " + protectedSheet.IsProtected);

            // Save the unprotected workbook
            protectedWb.Save("UnprotectedWorksheet.xlsx");
        }
    }
}