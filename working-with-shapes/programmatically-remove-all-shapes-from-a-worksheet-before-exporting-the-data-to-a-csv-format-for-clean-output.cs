using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RemoveShapesAndExportCsv
{
    static void Main()
    {
        // Load the workbook (replace with your source file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and remove every drawing object (shapes, charts, pictures, etc.)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Removes all drawing objects in the current worksheet
            sheet.RemoveAllDrawingObjects();
        }

        // Export the cleaned workbook to CSV format
        // SaveFormat.Csv writes the first worksheet by default; to export all sheets you could loop and save each.
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}