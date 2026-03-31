using System;
using Aspose.Cells;

class DuplicateWorksheets
{
    static void Main()
    {
        // Load the existing workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Names of worksheets that need to be duplicated
        string[] sheetsToCopy = { "Sheet1", "Data" };

        foreach (string sheetName in sheetsToCopy)
        {
            // Verify the source worksheet exists
            if (workbook.Worksheets[sheetName] != null)
            {
                // Add a copy of the worksheet; this preserves formatting, formulas, and references
                int copiedIndex = workbook.Worksheets.AddCopy(sheetName);
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

                // Optionally rename the copied worksheet to avoid name clashes
                copiedSheet.Name = sheetName + "_Copy";
            }
        }

        // Save the workbook with the duplicated worksheets
        string outputFile = "output.xlsx";
        workbook.Save(outputFile);
    }
}