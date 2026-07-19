// Title: Validate Empty Cells in Chart Data Range with Aspose.Cells for .NET (C#)
// Description: C# example that scans a worksheet range for blank cells, flags any gaps, sets PlotEmptyCellsType accordingly, creates a column chart and saves the workbook.
// Keywords: Aspose.Cells empty cells chart | C# validate chart data range | PlotEmptyCellsType Aspose.Cells | detect blank cells Excel range | Aspose.Cells column chart example | chart data validation .NET | Aspose.Cells range iteration
// Common Searches: how to check for blank cells before creating a chart in Aspose.Cells | Aspose.Cells C# PlotEmptyCellsType Zero vs NotPlotted | validate Excel range for null values using Aspose.Cells | C# example for chart data validation with Aspose.Cells | skip empty cells in Aspose.Cells chart series
// Developer Intent: Identify and handle empty cells in a chart’s source range so the chart renders correctly or applies a chosen empty‑cell plotting rule.
// Use Cases: Scanning a numeric range to ensure all required values are present before binding it to a chart series. | Automatically switching the chart’s PlotEmptyCellsType to Zero when blanks exist, otherwise leaving them unplotted. | Providing a runtime warning or log entry when empty cells are detected in chart data. | Integrating range validation into a larger Excel report generation workflow.
// AI Prompts: Generate C# code with Aspose.Cells that validates a given range for null or empty values and returns a boolean flag. | Show how to set Chart.PlotEmptyCellsType to Zero or NotPlotted based on the result of a range‑validation check. | Create a reusable method in Aspose.Cells that logs a warning if any blank cells are found before adding a chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// C# example that scans a worksheet range for blank cells, flags any gaps, sets PlotEmptyCellsType accordingly, creates a column chart and saves the workbook.
class ValidateChartData
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with an intentional empty cell (B3)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            // B3 left empty
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Define the ranges for the series values and categories
            string valueRange = "B2:B4";
            string categoryRange = "A2:A4";

            // Validate that the value range does not contain empty cells
            bool containsEmpty = false;
            AsposeRange rangeToCheck = worksheet.Cells.CreateRange(valueRange);
            foreach (Cell cell in rangeToCheck)
            {
                // Aspose.Cells does not expose IsBlank in some versions; use Value check instead
                if (cell.Value == null || string.IsNullOrEmpty(cell.StringValue))
                {
                    containsEmpty = true;
                    break;
                }
            }

            if (containsEmpty)
            {
                Console.WriteLine("Empty cells detected in the data range. The chart will plot them as zeros.");
            }

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add(valueRange, true);
            chart.NSeries.CategoryData = categoryRange;

            // Configure how empty cells are plotted based on validation result
            chart.PlotEmptyCellsType = containsEmpty ? PlotEmptyCellsType.Zero : PlotEmptyCellsType.NotPlotted;

            // Determine output file path and ensure the directory exists
            string outputPath = "ValidatedChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
