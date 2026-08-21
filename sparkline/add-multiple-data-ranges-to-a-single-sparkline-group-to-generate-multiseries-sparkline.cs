// Title: C# – Add Multiple Data Ranges to a Single SparklineGroup for a Multi‑Series Line Sparkline (Aspose.Cells)
// Description: This Aspose.Cells for .NET example creates a workbook, fills three columns with numeric data, defines a vertical sparkline location (D1:D3), adds a line SparklineGroup, and uses SparklineGroup.ResetRanges with a comma‑separated range (A1:A5,B1:B5,C1:C5) to bind three series to one sparkline. The code also sets a custom series color and enables high/low point markers before saving the file as MultiSeriesSparkline.xlsx.
// Keywords: Aspose.Cells C# sparkline multiple ranges | multi‑series sparkline .NET | SparklineGroup ResetRanges example | vertical sparkline group Aspose.Cells | custom series color sparkline | Aspose.Cells SparklineGroup API | C# Excel sparkline tutorial
// Common Searches: how to add several data ranges to a SparklineGroup in Aspose.Cells | Aspose.Cells multi‑series sparkline from columns | ResetRanges vertical sparkline .NET example | set series color for multi‑series sparkline Aspose.Cells | C# create line sparkline with multiple series
// Developer Intent: Bind multiple column‑based data series to one SparklineGroup to generate a single multi‑series line sparkline and customize its visual style.
// Use Cases: Show compact trend lines for sales, profit, and quantity side‑by‑side in a dashboard worksheet. | Compare monthly performance of different product categories using vertical sparklines that pull data from separate columns. | Highlight key points in a financial report by applying a custom series color and high/low markers to a multi‑series sparkline.
// AI Prompts: Generate C# code with Aspose.Cells that adds three column ranges to a SparklineGroup, sets the sparkline orientation to vertical, and applies a custom series color and high/low markers. | Explain each parameter of SparklineGroup.ResetRanges for creating multi‑series sparklines and demonstrate how to pass a comma‑separated range string. | Provide a step‑by‑step guide to build a multi‑series line sparkline, customize its appearance, and save the workbook as an .xlsx file using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells for .NET example creates a workbook, fills three columns with numeric data, defines a vertical sparkline location (D1:D3), adds a line SparklineGroup, and uses SparklineGroup.ResetRanges with a comma‑separated range (A1:A5,B1:B5,C1:C5) to bind three series to one sparkline. The code also sets a custom series color and enables high/low point markers before saving the file as MultiSeriesSparkline.xlsx.
class MultiSeriesSparklineDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series (each in its own column)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);          // Series 1 in column A
                sheet.Cells[i, 1].PutValue((i + 1) * 2);   // Series 2 in column B
                sheet.Cells[i, 2].PutValue((i + 1) * 3);   // Series 3 in column C
            }

            // Define the location range where the sparklines will be placed (D1:D3)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 3,
                EndColumn = 3
            };

            // Add a sparkline group of type Line
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add multiple data ranges (comma‑separated) to the same group.
            // Set isVertical = true because each series is stored in a column.
            string multiDataRange = "A1:A5,B1:B5,C1:C5";
            group.ResetRanges(multiDataRange, true, location);

            // Optional: customize the appearance of the sparkline group
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue;
            group.SeriesColor = seriesColor;
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            // Save the workbook
            workbook.Save("MultiSeriesSparkline.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
