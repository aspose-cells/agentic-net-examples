using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace SparklineXLTMExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 0, columns A-D)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group of type Line, using the data range A1:D1,
            // not vertical (plot by column), and place it in the defined location
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                sheet.Name + "!A1:D1",
                false,
                sparklineLocation);

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group (the same data range, positioned at row 0, column 4)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Optional: customize appearance (e.g., set series color)
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = System.Drawing.Color.Orange;
            group.SeriesColor = seriesColor;

            // Save the workbook as an Excel Macro‑Enabled Template (XLTM)
            workbook.Save("SparklineTemplate.xltm", SaveFormat.Xltm);
        }
    }
}