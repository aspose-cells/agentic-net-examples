// Title: C# – Validate Empty Cells in Chart Data Range with Aspose.Cells
// Description: This C# example demonstrates scanning a worksheet range for null values, aborting chart creation when empty cells are found, and generating a column chart only if the data is complete. It covers workbook creation, defining value and category ranges, checking for empty cells, setting PlotEmptyCellsType to NotPlotted, and saving the file.
// Keywords: Aspose.Cells C# chart validation | check empty cells before chart | column chart Aspose.Cells | prevent chart errors Aspose | worksheet range null detection | PlotEmptyCellsType NotPlotted
// Common Searches: Aspose.Cells check for empty cells before creating chart | C# validate chart data range Aspose | how to abort chart creation when data missing Aspose.Cells | detect null values in Excel range using Aspose.Cells | set PlotEmptyCellsType after data validation
// Developer Intent: Verify that the data range used for a chart contains no empty cells and skip chart creation if any are found.
// Use Cases: Automated report generators that must avoid charts with incomplete data | Data quality checks in Excel export pipelines using Aspose.Cells | Dynamic dashboards that only display charts when all required values are present
// AI Prompts: Generate a function that returns the address of the first empty cell in a specified range using Aspose.Cells. | Create a logger that records every empty cell found before building multiple charts in a workbook. | Suggest an alternative method using Worksheet.Cells.MaxDataColumn/Row to identify empty cells within a range.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// This C# example demonstrates scanning a worksheet range for null values, aborting chart creation when empty cells are found, and generating a column chart only if the data is complete. It covers workbook creation, defining value and category ranges, checking for empty cells, setting PlotEmptyCellsType to NotPlotted, and saving the file.
class ValidateChartData
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (some cells may be empty)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            // B3 is intentionally left empty
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Define the data ranges for values and categories
            string valueRange = "B2:B4";
            string categoryRange = "A2:A4";

            // Validate that the value range does not contain empty cells
            if (ContainsEmptyCells(sheet, valueRange))
            {
                Console.WriteLine("The data range contains empty cells. Chart creation aborted.");
            }
            else
            {
                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the series and category data
                chart.NSeries.Add(valueRange, true);
                chart.NSeries.CategoryData = categoryRange;

                // Explicitly set how empty cells should be handled (not needed after validation)
                chart.PlotEmptyCellsType = PlotEmptyCellsType.NotPlotted;

                Console.WriteLine("Chart created successfully.");
            }

            // Save the workbook
            string outputPath = "ValidatedChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to check for any empty cells within a given range address
    static bool ContainsEmptyCells(Worksheet sheet, string rangeAddress)
    {
        AsposeRange range = sheet.Cells.CreateRange(rangeAddress);
        foreach (Cell cell in range)
        {
            // A cell is considered empty if its value is null
            if (cell.Value == null)
                return true; // Empty cell found
        }
        return false; // No empty cells
    }
}
