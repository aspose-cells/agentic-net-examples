using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace SparklineLightingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group of type Line with the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group (the same range, placed at row 0, column 4)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // NOTE:
            // Aspose.Cells does not expose a direct lighting property for sparklines.
            // Lighting effects such as LightRigType.Soft are applicable to 3‑D shapes
            // (e.g., chart series) via the Format3D property. Since sparklines are
            // rendered as 2‑D objects, they do not support lighting adjustments.
            // Therefore, no code is required to set lighting for a sparkline.

            // Save the workbook as an Excel 2010 file
            workbook.Save("SparklineWithSoftLighting.xlsx", SaveFormat.Xlsx);
        }
    }
}