using System;
using Aspose.Cells;

class WorksheetPasswordProtectionDemo
{
    static void Main()
    {
        // Create a new workbook and add some data to the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");

        // Protect the worksheet with a password and all protection types
        sheet.Protect(ProtectionType.All, "mySecretPassword", null);

        // Save the workbook with the protected worksheet
        workbook.Save("ProtectedWorksheet.xlsx");

        // Load the workbook (worksheet protection does not require a load password)
        Workbook loadedWorkbook = new Workbook("ProtectedWorksheet.xlsx");
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // Display the protection status
        Console.WriteLine("Worksheet is protected: " + loadedSheet.IsProtected);

        // Unprotect the worksheet using the correct password
        loadedSheet.Unprotect("mySecretPassword");

        // Save the workbook after removing protection
        loadedWorkbook.Save("UnprotectedWorksheet.xlsx");
    }
}