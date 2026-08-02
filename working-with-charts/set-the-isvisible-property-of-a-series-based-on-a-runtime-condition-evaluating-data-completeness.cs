// Title: Conditionally hide a chart series in Aspose.Cells for .NET using IsFiltered
// Description: Shows how to scan a worksheet range for blank or whitespace cells, decide if the data is complete, and set the chart series IsFiltered property to hide or show the series before saving the workbook.
// Keywords: Aspose.Cells chart series visibility | C# IsFiltered property | hide chart series Aspose.Cells | check empty cells Excel chart | conditional chart series .NET | Aspose.Cells runtime condition | CellArea iteration | Excel chart data validation | Aspose.Cells example | C# Excel chart hide series
// Common Searches: Aspose.Cells hide series when data missing | C# set chart series IsFiltered | check for blank cells before chart Aspose.Cells | conditional visibility of chart series .NET | how to hide column chart series programmatically | Aspose.Cells chart series visibility based on data | C# iterate over CellArea Aspose.Cells | Excel chart hide incomplete series
// Developer Intent: Programmatically evaluate a data range and hide the associated chart series if any cell is empty or contains only whitespace.
// Use Cases: Automated report generation that excludes incomplete series to avoid misleading charts. | Dashboard creation where series appear only when all required data points are present. | Data‑quality checks that automatically filter out partial data before visualizing. | Dynamic Excel export where series visibility adapts to runtime data conditions.
// AI Prompts: Generate C# code with Aspose.Cells that scans a CellArea for null or whitespace values and sets the chart series IsFiltered flag accordingly. | Show an example of toggling a column chart series visibility based on completeness of the source range using Aspose.Cells. | Explain the steps to validate data completeness and hide a chart series in Aspose.Cells, including CellArea definition and IsFiltered usage. | Provide a snippet that hides a chart series when any cell in the series data range is empty, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    // Shows how to scan a worksheet range for blank or whitespace cells, decide if the data is complete, and set the chart series IsFiltered property to hide or show the series before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with a missing value to simulate incompleteness
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["A5"].PutValue("D");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue("");   // Missing value
                worksheet.Cells["B5"].PutValue(40);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the series data range
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Determine if the series data is complete (no empty cells)
                bool isComplete = true;
                CellArea dataArea = new CellArea
                {
                    StartRow = 1, // B2 (zero‑based index)
                    EndRow = 4,   // B5
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Iterate through each cell in the range and check for blank/empty values
                for (int row = dataArea.StartRow; row <= dataArea.EndRow; row++)
                {
                    Cell cell = worksheet.Cells[row, dataArea.StartColumn];
                    // Treat null, empty string, or whitespace as missing data
                    if (cell.Value == null || string.IsNullOrWhiteSpace(cell.StringValue))
                    {
                        isComplete = false;
                        break;
                    }
                }

                // Hide the series if data is incomplete; otherwise, show it
                chart.NSeries[0].IsFiltered = !isComplete;

                // Save the workbook
                string outputPath = "SeriesVisibilityBasedOnDataCompleteness.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
