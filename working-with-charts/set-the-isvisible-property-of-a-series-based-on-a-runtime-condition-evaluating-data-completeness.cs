// Title: Conditionally Hide or Show a Chart Series at Runtime with Aspose.Cells for .NET
// Description: C# example that creates a workbook, builds a column chart, checks a series range for blank cells, and toggles the series visibility using the IsFiltered property based on data completeness.
// Keywords: Aspose.Cells chart series visibility | IsFiltered property C# | hide chart series runtime | check data completeness Aspose.Cells | conditional chart series Aspose.Cells | Excel chart series filter .NET | dynamic series visibility
// Common Searches: how to hide a chart series in Aspose.Cells when data is missing | set IsFiltered for a series based on a condition | evaluate cell range before displaying chart series Aspose.Cells | conditional chart series visibility .NET | filter out incomplete series in Excel chart using Aspose
// Developer Intent: Determine at execution time whether a chart series should be displayed by scanning its source cells for blanks and setting the series' IsFiltered flag accordingly.
// Use Cases: Automatically exclude series with incomplete data from generated reports. | Create dashboards that only show fully populated series based on user selections. | Prevent misleading charts by hiding series containing null or empty values.
// AI Prompts: Generate C# code with Aspose.Cells that hides a chart series when any cell in its data range is empty by using the IsFiltered property. | Show how to iterate over a CellArea to verify data completeness and then toggle series visibility with a boolean flag. | Explain the role of the IsFiltered property for conditional chart series display in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    // C# example that creates a workbook, builds a column chart, checks a series range for blank cells, and toggles the series visibility using the IsFiltered property based on data completeness.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (some cells intentionally left blank to simulate incompleteness)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["A5"].PutValue("D");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                // B4 left blank
                worksheet.Cells["B5"].PutValue(40);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Evaluate data completeness for the series range B2:B5
                bool isComplete = true;
                CellArea range = new CellArea { StartRow = 1, EndRow = 4, StartColumn = 1, EndColumn = 1 }; // B2:B5
                for (int row = range.StartRow; row <= range.EndRow; row++)
                {
                    Cell cell = worksheet.Cells[row, range.StartColumn];
                    // Consider a cell blank if its value is null or an empty string
                    if (cell.Value == null || string.IsNullOrEmpty(cell.StringValue))
                    {
                        isComplete = false;
                        break;
                    }
                }

                // Set series visibility based on completeness.
                // In Aspose.Cells, hiding a series is done via the IsFiltered property.
                // True means the series is filtered out (not displayed).
                chart.NSeries[0].IsFiltered = !isComplete;

                // Define output file path
                string outputPath = "SeriesVisibilityBasedOnDataCompleteness.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
