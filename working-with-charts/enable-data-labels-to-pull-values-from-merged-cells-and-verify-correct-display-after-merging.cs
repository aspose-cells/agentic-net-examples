// Title: Aspose.Cells for .NET – Use Merged Cells as Chart Data Labels (C# Example)
// Description: This C# demo creates a workbook, merges cells C2:C3, adds a column chart, and configures the series so data labels pull custom text from the merged range using ShowCellRange and LinkedSource. It verifies the merged value and saves the file as DataLabelsFromMergedCellsDemo.xlsx.
// Keywords: Aspose.Cells | C# chart data labels | merged cells | ShowCellRange | LinkedSource | column chart | Aspose.Cells for .NET example | chart label from cell range | Excel automation | DataLabelsFromMergedCellsDemo
// Common Searches: Aspose.Cells merged cells data label | ShowCellRange property C# | LinkedSource chart label Aspose | use merged cell as chart label .NET | Aspose.Cells column chart custom labels
// Developer Intent: Display custom text from a merged cell range as data labels on a chart.
// Use Cases: Generate a column chart where each point shows a unit label stored in a merged cell. | Consolidate identical labels across multiple series by merging cells and linking them to the chart. | Validate that merged cell values appear correctly in exported Excel files.
// AI Prompts: Write C# code with Aspose.Cells that merges C2:C3 and sets them as the LinkedSource for chart data labels. | Explain the interaction between ShowCellRange and LinkedSource when using merged cells for chart labels. | Troubleshoot why data labels might ignore a merged cell source in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // This C# demo creates a workbook, merges cells C2:C3, adds a column chart, and configures the series so data labels pull custom text from the merged range using ShowCellRange and LinkedSource. It verifies the merged value and saves the file as DataLabelsFromMergedCellsDemo.xlsx.
    public class DataLabelsFromMergedCellsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data
                // Column A – Category
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("A");
                cells["A3"].PutValue("B");

                // Column B – Values for the chart
                cells["B1"].PutValue("Value");
                cells["B2"].PutValue(100);
                cells["B3"].PutValue(200);

                // Column C – Labels that will be merged
                cells["C1"].PutValue("Label");
                cells["C2"].PutValue("100 units");
                cells["C3"].PutValue("200 units");

                // Merge the label cells C2:C3 into a single cell (C2)
                // After merging, the value of the upper‑left cell (C2) will be used for both data points
                cells.Merge(1, 2, 2, 1); // firstRow=1 (C2), firstColumn=2 (C), totalRows=2, totalColumns=1

                // Verify that the merge succeeded
                Console.WriteLine($"Cell C2 IsMerged: {cells["C2"].IsMerged}");
                Console.WriteLine($"Merged cell value (C2): {cells["C2"].StringValue}");

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and the category axis
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Configure data labels to pull values from the merged cell range
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;          // Show the chart's own value
                series.DataLabels.ShowCellRange = true;      // Enable pulling from a cell range
                series.DataLabels.LinkedSource = "C2:C3";    // Range that contains the merged cells
                series.DataLabels.Font.Color = Color.Blue;  // Optional styling

                // Save the workbook
                string outputPath = "DataLabelsFromMergedCellsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsFromMergedCellsDemo.Run();
        }
    }
}
