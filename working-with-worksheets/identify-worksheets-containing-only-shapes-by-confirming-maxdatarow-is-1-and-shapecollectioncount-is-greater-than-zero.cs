using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class IdentifyShapeOnlySheets
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow == -1 means the worksheet has no cell data
            bool hasNoCellData = sheet.Cells.MaxDataRow == -1;

            // Shapes.Count > 0 means the worksheet contains drawing objects
            bool hasShapes = sheet.Shapes.Count > 0;

            // If both conditions are true, the sheet contains only shapes
            if (hasNoCellData && hasShapes)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains only shapes.");
            }
        }

        // Save the workbook (optional, can be omitted if only reading)
        workbook.Save("output.xlsx");
    }
}