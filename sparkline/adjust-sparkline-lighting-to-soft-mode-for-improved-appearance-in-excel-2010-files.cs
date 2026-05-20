using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class AdjustSparklineLighting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location where the sparkline will be placed
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group (Line type) with the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // NOTE:
        // Aspose.Cells does not expose a direct property to modify the lighting of a sparkline.
        // Lighting (e.g., LightRigType.Soft) is applicable to 3‑D shape formatting, which is
        // available for chart series and other shapes, but not for sparkline groups.
        // Therefore, the sparkline lighting cannot be set to 'Soft' via the current API.

        // Save the workbook as an Excel 2010 file
        workbook.Save("SparklineWithSoftLighting.xlsx", SaveFormat.Xlsx);
    }
}