// Title: Aspose.Cells C# – Use Merged Cells as Chart Data Labels and Verify Output
// Description: Shows how to merge cells, link the merged range as a data‑label source for a column chart, style the labels, save the workbook, and print merged‑cell details for verification with Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart merged cells | C# data labels linked source | ShowCellRange Aspose.Cells | merged cell value as chart label | verify merged cell in Aspose.Cells | Aspose.Cells .NET chart example | DataLabelsFromMergedCellsDemo
// Common Searches: Aspose.Cells chart data labels from merged cells | C# Aspose.Cells ShowCellRange example | How to use LinkedSource with merged cells in Aspose.Cells | Display merged cell value as chart label Aspose.Cells | Verify merged cell information in Aspose.Cells workbook
// Developer Intent: The developer wants to configure a chart so its data labels are taken from a merged cell range and confirm that the correct label value appears.
// Use Cases: Create a column chart where each series label is sourced from a merged range (e.g., C2:C4) using ShowCellRange and LinkedSource. | Merge cells to represent a single descriptive label and ensure the chart reads the top‑left cell value for that label. | Save the workbook and output the merged cell address, merge flag, and string value to validate label accuracy.
// AI Prompts: Generate C# code with Aspose.Cells that merges cells, links them as data‑label sources for a chart, and prints merged‑cell details for verification. | Explain the interaction between ShowCellRange and LinkedSource when pulling labels from merged cells in Aspose.Cells charts. | Provide troubleshooting steps if chart data labels do not display values from merged cells as expected.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to merge cells, link the merged range as a data‑label source for a column chart, style the labels, save the workbook, and print merged‑cell details for verification with Aspose.Cells for .NET.
    public class DataLabelsFromMergedCellsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // Prepare sample data for the chart
                // -------------------------------------------------
                // Category column
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("Item 1");
                cells["A3"].PutValue("Item 2");
                cells["A4"].PutValue("Item 3");

                // Values column (numeric data)
                cells["B1"].PutValue("Value");
                cells["B2"].PutValue(120);
                cells["B3"].PutValue(80);
                cells["B4"].PutValue(150);

                // Labels column – we will merge some cells here
                cells["C1"].PutValue("Label");
                cells["C2"].PutValue("High");
                cells["C3"].PutValue("Medium");
                cells["C4"].PutValue("Low");

                // Merge label cells C2:C3 to simulate a merged label
                // After merging, the value of the merged cell is taken from the top‑left cell (C2)
                cells.Merge(1, 2, 2, 1); // rows 1‑2 (zero‑based), column 2 (C)

                // -------------------------------------------------
                // Create a column chart
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the series (values) and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Configure data labels to pull values from the merged cells
                // -------------------------------------------------
                var series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;               // show the numeric value
                series.DataLabels.ShowCellRange = true;           // enable using cell range as label source
                series.DataLabels.LinkedSource = "C2:C4";         // range that includes the merged cell

                // Optional: style the data labels
                series.DataLabels.Font.Color = Color.Blue;
                series.DataLabels.Font.IsBold = true;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string filePath = "DataLabelsFromMergedCellsDemo.xlsx";

                // Ensure the directory exists (in case a relative path is used)
                string directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(filePath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(filePath)}");

                // -------------------------------------------------
                // Verification – output merged cell information
                // -------------------------------------------------
                Cell mergedCell = cells["C2"]; // top‑left cell of the merged area
                Console.WriteLine($"Merged cell address: {mergedCell.Name}");
                Console.WriteLine($"Is merged: {mergedCell.IsMerged}");
                Console.WriteLine($"Value used for data label: {mergedCell.StringValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsFromMergedCellsDemo.Run();
        }
    }
}
