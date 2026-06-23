using System;
using System.IO;
using Aspose.Cells;

class ProtectWorksheets
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path where the protected workbook will be saved
        string outputPath = "output_protected.xlsx";

        // Load the workbook (lifecycle rule)
        Workbook workbook = new Workbook(inputPath);

        // Derive a base password from the file name (without extension)
        string basePassword = Path.GetFileNameWithoutExtension(inputPath);

        // Protect each worksheet with a unique password
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Create a unique password per worksheet (e.g., "MyFile_Sheet0")
            string sheetPassword = $"{basePassword}_Sheet{sheet.Index}";

            // Protect the worksheet with all protection types and the unique password
            sheet.Protect(ProtectionType.All, sheetPassword, null);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save(outputPath);
    }
}