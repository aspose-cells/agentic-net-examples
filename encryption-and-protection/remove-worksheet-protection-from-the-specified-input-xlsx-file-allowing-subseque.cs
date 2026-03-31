using Aspose.Cells;
using System;

class RemoveWorksheetProtection
{
    static void Main()
    {
        // Path to the input workbook that has worksheet protection
        string inputPath = "protected_input.xlsx";

        // Load the workbook (uses the provided load rule)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and remove protection
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Unprotect the worksheet. If a password is required, provide it here:
            // sheet.Unprotect("yourPassword");
            sheet.Unprotect();
        }

        // Save the workbook after removing protection (uses the provided save rule)
        string outputPath = "unprotected_output.xlsx";
        workbook.Save(outputPath);
    }
}