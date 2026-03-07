using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace SparklineXlsmDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 1, columns A-D)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,   // Row index is zero‑based
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group of type Line, using the data range A1:D1,
            // not vertical (plot by column), and place the sparkline in E1
            int groupIdx = sheet.SparklineGroups.Add(
                SparklineType.Line,
                sheet.Name + "!A1:D1",
                false,
                sparklineLocation);

            // Retrieve the created group
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline to the group (data range A1:D1, row 0, column 4)
            // This call is optional because the Add method above already created the sparkline,
            // but it demonstrates the explicit addition.
            int sparklineIdx = group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);
            Sparkline sparkline = group.Sparklines[sparklineIdx];

            // Optional: customize appearance (e.g., series color)
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = System.Drawing.Color.Orange;
            group.SeriesColor = seriesColor;

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save("SparklineExample.xlsm", SaveFormat.Xlsm);
        }
    }
}