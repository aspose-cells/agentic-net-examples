// Title: Move an Aspose.Cells chart to start at cell D5 by setting UpperLeftRow and UpperLeftColumn in C#
// Common Searches: Aspose.Cells C# set chart UpperLeftRow to position chart at row 5 | C# Aspose.Cells place chart in cell D5 using ChartObject properties | How to change chart location to a specific cell with Aspose.Cells .NET | Adjust Excel chart upper‑left corner to column D row 5 in Aspose.Cells | Set chart UpperLeftColumn and UpperLeftRow for precise placement in Aspose.Cells
// Tags: Aspose.Cells chart UpperLeftRow property usage | C# ChartObject UpperLeftColumn for Excel chart placement | move chart to cell D5 with Aspose.Cells | precise chart positioning in Excel using Aspose.Cells .NET | adjust chart location without resizing Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartPositionDemo
{
    // The example creates a workbook, adds sample data, inserts a column chart, and then moves the chart's upper‑left corner to cell D5 by assigning ChartObject.UpperLeftColumn = 3 and UpperLeftRow = 4, preserving the original size before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            string outputPath = "ChartMovedToD5.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart (initial position is arbitrary)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Move the chart so that its upper‑left corner starts at cell D5
            // Aspose.Cells uses zero‑based indices: D = column 3, row 5 = row index 4
            chart.ChartObject.UpperLeftColumn = 3; // Column D
            chart.ChartObject.UpperLeftRow = 4;    // Row 5

            // Optionally, adjust the lower‑right corner to keep the size unchanged
            // (If you want to keep the original size, you can also set LowerRightRow/Column)
            // Here we simply keep the existing size, so no changes are needed.

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"Chart upper‑left corner is now at column {chart.ChartObject.UpperLeftColumn} (D), row {chart.ChartObject.UpperLeftRow} (5).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
    }
}
