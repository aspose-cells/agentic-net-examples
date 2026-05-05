using System;
using Aspose.Cells;

class RemoveWorksheetProtection
{
    static void Main()
    {
        string inputPath = "protected.xlsx";
        string password = "password123";

        Workbook workbook = new Workbook(inputPath);

        // Unprotect the workbook (structure and windows) with the password
        workbook.Unprotect(password);

        // Unprotect each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.IsProtected)
                sheet.Unprotect();
        }

        string outputPath = "unprotected.xlsx";
        workbook.Save(outputPath);
    }
}