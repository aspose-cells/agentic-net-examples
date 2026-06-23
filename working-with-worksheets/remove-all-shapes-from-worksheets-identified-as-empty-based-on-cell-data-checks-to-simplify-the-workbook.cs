using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RemoveShapesFromEmptySheets
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.IgnoreUselessShapes = true; // optional, improves loading performance
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet contains any data.
            // MaxDataRow and MaxDataColumn are -1 when the sheet is completely empty.
            if (sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0)
            {
                // The worksheet is empty – remove all drawing objects (shapes) from it.
                sheet.RemoveAllDrawingObjects();
            }
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}