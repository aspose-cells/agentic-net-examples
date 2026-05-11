using System;
using Aspose.Cells;

class CopySpecificWorksheets
{
    static void Main()
    {
        // Paths to the source and target workbooks
        string sourcePath = "source.xlsx";
        string targetPath = "target.xlsx";

        // Load the source workbook from file
        Workbook sourceWorkbook = new Workbook(sourcePath);

        // Create a new (empty) target workbook
        Workbook targetWorkbook = new Workbook();

        // Remove the default worksheet that Aspose.Cells creates
        if (targetWorkbook.Worksheets.Count > 0)
        {
            targetWorkbook.Worksheets.Clear();
        }

        // List of worksheet names that need to be copied from the source workbook
        string[] sheetsToCopy = { "Sheet1", "Data", "Summary" };

        foreach (string sheetName in sheetsToCopy)
        {
            // Verify that the source workbook actually contains the worksheet
            Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];
            if (sourceSheet == null)
                continue; // Skip if the sheet does not exist

            // Add a new worksheet to the target workbook with the same name
            Worksheet targetSheet = targetWorkbook.Worksheets.Add(sheetName);

            // Copy contents and formats from the source worksheet to the target worksheet
            targetSheet.Copy(sourceSheet);
        }

        // Save the target workbook to the specified file
        targetWorkbook.Save(targetPath);
    }
}