// Title: Add a top legend to a column chart and freeze the legend rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to add a top legend to a column chart and freeze its rows using Aspose.Cells | Aspose.Cells freeze panes for rows that contain a chart legend in .NET | How to keep chart legend visible while scrolling in an Excel file with Aspose.Cells | Set legend position and freeze top rows in Aspose.Cells chart example | Create column chart with legend and freeze first rows in C# Aspose.Cells
// Tags: add top legend to column chart Aspose.Cells C# | freeze worksheet rows containing chart legend Aspose.Cells | configure legend size and font Aspose.Cells chart | column chart with frozen legend rows Aspose.Cells | save workbook with chart and frozen panes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendFreezeDemo
{
    // The example creates a new workbook, fills it with sample data, adds a column chart, positions a custom-sized top legend, freezes the first five rows so the legend stays visible while scrolling, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for the chart
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["A3"].PutValue("B");
                cells["A4"].PutValue("C");
                cells["B2"].PutValue(10);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(30);

                // 3. Add a column chart to the worksheet
                // Position the chart so that its legend will be in the top rows (rows 2‑4)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 2, 0, 15, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // 4. Configure the legend
                Legend legend = chart.Legend;
                legend.Position = LegendPositionType.Top;   // Place legend above the plot area
                legend.IsAutomaticSize = false;            // Allow manual sizing
                legend.Width = 400;                        // Width in points
                legend.Height = 30;                        // Height in points
                legend.Font.Size = 11;
                legend.Font.IsBold = true;
                // legend.IsOverlay property does not exist; removed.

                // 5. Freeze the rows that contain the legend.
                // Freeze the first 5 rows (0‑based index 4) and no columns.
                sheet.FreezePanes(5, 0, 5, 0);

                // 6. Save the workbook
                string outputPath = "ChartWithLegendAndFrozenRows.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Report any errors without hiding them
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
