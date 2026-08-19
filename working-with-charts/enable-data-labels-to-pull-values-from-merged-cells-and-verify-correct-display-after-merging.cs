// Title: Aspose.Cells for .NET – Pull chart data labels from merged cells and verify output
// Description: Demonstrates how to merge cells, assign custom label text, create a column chart, enable ShowCellRange, link the series data labels to the merged ranges via LinkedSource, style the labels, print merged‑cell verification details, and save the workbook.
// Keywords: Aspose.Cells C# chart data labels | merged cells chart labels | ShowCellRange Aspose.Cells | LinkedSource merged range | column chart custom labels | verify merged cells Aspose | Excel automation Aspose.Cells
// Common Searches: Aspose.Cells chart label from merged cell | C# link data labels to merged range | ShowCellRange property example | How to use LinkedSource with merged cells | Validate merged cells before charting
// Developer Intent: Connect chart data labels to values stored in merged cells and confirm they render correctly.
// Use Cases: Generate reports where each column label displays a formatted string from a vertically merged cell. | Automate Excel workbooks that keep label text separate from data values, then pull it into chart labels. | Programmatically check IsMerged and cell content before linking to ensure label accuracy.
// AI Prompts: Write C# code with Aspose.Cells that merges cells, sets custom label text, creates a column chart, enables ShowCellRange, and links data labels using LinkedSource. | Explain the interaction between Series.DataLabels.ShowCellRange, LinkedSource, and merged cells in Aspose.Cells. | Provide a try‑catch block that validates merged‑cell status and logs any issues before saving the workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMergedDataLabelsDemo
{
    // Demonstrates how to merge cells, assign custom label text, create a column chart, enable ShowCellRange, link the series data labels to the merged ranges via LinkedSource, style the labels, print merged‑cell verification details, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- Populate source data ----------
                // Categories
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("A");
                cells["A3"].PutValue("B");

                // Values for the chart
                cells["B1"].PutValue("Value");
                cells["B2"].PutValue(100);
                cells["B3"].PutValue(200);

                // Create merged cells that will serve as data label sources
                // Merge C2:C3 vertically and put a label value in the merged cell
                cells.Merge(1, 2, 2, 1); // rows 1‑2 (zero‑based), column 2 (C)
                cells["C2"].PutValue("100 units");

                // Merge D2:D3 for the second data point
                cells.Merge(1, 3, 2, 1);
                cells["D2"].PutValue("200 units");

                // ---------- Create a column chart ----------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B3", true);          // Values
                chart.NSeries.CategoryData = "A2:A3";      // Categories

                // ---------- Configure data labels to use merged cells ----------
                // Use the first series (index 0)
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;               // Show the numeric value
                series.DataLabels.ShowCellRange = true;           // Enable pulling from cell range
                // Link the data labels to the merged ranges C2:C3 and D2:D3
                series.DataLabels.LinkedSource = "C2:D3";

                // Optional: style the data labels
                series.DataLabels.Font.Color = Color.Blue;

                // ---------- Verify merged cells ----------
                Console.WriteLine("Verification of merged cells used for data labels:");
                Console.WriteLine($"C2 IsMerged: {cells["C2"].IsMerged}");
                Console.WriteLine($"C2 Value   : {cells["C2"].StringValue}");
                Console.WriteLine($"D2 IsMerged: {cells["D2"].IsMerged}");
                Console.WriteLine($"D2 Value   : {cells["D2"].StringValue}");

                // Save the workbook
                string outputPath = "MergedDataLabelsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
