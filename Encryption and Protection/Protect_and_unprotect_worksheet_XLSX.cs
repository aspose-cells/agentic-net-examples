using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password and all protection types
            sheet.Protect(ProtectionType.All, "myPassword123", null);
            Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

            // Save the protected workbook
            workbook.Save("ProtectedWorksheet.xlsx");

            // Load the saved workbook
            Workbook loadedWorkbook = new Workbook("ProtectedWorksheet.xlsx");
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Loaded worksheet protected: " + loadedSheet.IsProtected);

            // Unprotect the worksheet using the correct password
            loadedSheet.Unprotect("myPassword123");
            Console.WriteLine("Worksheet unprotected: " + !loadedSheet.IsProtected);

            // Save the unprotected workbook
            loadedWorkbook.Save("UnprotectedWorksheet.xlsx");
        }
    }
}