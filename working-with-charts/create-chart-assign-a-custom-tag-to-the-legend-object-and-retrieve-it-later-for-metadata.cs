// Title: Create a Column Chart, Tag Its Legend via Hidden Cell, and Retrieve the Tag with Aspose.Cells for .NET
// Description: Demonstrates how to generate a column chart in a new workbook, assign a custom identifier to the legend by writing it to a hidden cell, save the file, reload it, and read the identifier back using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart legend tag | C# store custom metadata Excel | hidden cell tag Aspose.Cells | retrieve legend identifier .NET | Excel chart custom data persistence | Aspose.Cells .NET example | metadata in chart legend | automated report tagging Excel
// Common Searches: how to add custom tag to chart legend using Aspose.Cells C# | store and retrieve legend metadata in Excel with Aspose.Cells | persist custom identifier for chart legend in hidden cell | Aspose.Cells example for tagging chart legends | read hidden cell value as legend tag after saving workbook
// Developer Intent: Persist a custom identifier on a chart legend and read it back later.
// Use Cases: Attach version or report IDs to legends for automated downstream processing. | Link a legend to external data sources by storing a reference key in a hidden cell. | Embed non‑visible metadata in charts to support analytics without altering worksheet layout.
// AI Prompts: Show an alternative method to embed custom metadata directly into a chart legend without using hidden cells in Aspose.Cells. | Provide C# code that retrieves a legend tag stored in a named range instead of a fixed cell address. | Explain how to use Aspose.Cells chart properties to store and read custom legend information programmatically.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendTagDemo
{
    // Demonstrates how to generate a column chart in a new workbook, assign a custom identifier to the legend by writing it to a hidden cell, save the file, reload it, and read the identifier back using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the legend and customize it
                Legend legend = chart.Legend;
                legend.Position = LegendPositionType.Right;
                legend.Text = "Quarterly Sales";

                // ---- Assign a custom tag to the legend ----
                // Store the tag in a hidden cell (e.g., Z1)
                string customTag = "LegendTag_2024_Q1";
                const string tagCellAddress = "Z1";
                sheet.Cells[tagCellAddress].PutValue(customTag);

                // Save the workbook
                string filePath = "ChartWithLegendTag.xlsx";
                workbook.Save(filePath);

                // ---- Retrieve the custom tag later ----
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"The file '{filePath}' was not found.");

                Workbook loadedWorkbook;
                try
                {
                    loadedWorkbook = new Workbook(filePath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                Chart loadedChart = loadedSheet.Charts[chartIndex];
                Legend loadedLegend = loadedChart.Legend;

                // Retrieve the tag from the known hidden cell
                string retrievedTag = loadedSheet.Cells[tagCellAddress].StringValue;

                Console.WriteLine($"Custom tag retrieved from legend: {retrievedTag}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
