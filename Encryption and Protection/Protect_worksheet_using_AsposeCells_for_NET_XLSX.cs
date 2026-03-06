using System;
using Aspose.Cells;

class ProtectWorksheetDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with a password and all protection types
        // (type, password, oldPassword) overload is used as per the provided rule
        sheet.Protect(ProtectionType.All, "myPassword123", null);

        // Optional: adjust specific protection options
        Protection protection = sheet.Protection;
        protection.AllowEditingObject = false;      // Disallow editing objects
        protection.AllowEditingScenario = false;    // Disallow editing scenarios

        // Save the protected workbook as XLSX
        workbook.Save("ProtectedWorksheet.xlsx", SaveFormat.Xlsx);
    }
}