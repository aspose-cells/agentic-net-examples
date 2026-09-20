// Title: Generate a column pivot chart from an existing pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file, finds the first pivot table, adds a column chart linked to that pivot table on the same worksheet, sets the chart title to "Aggregated Sales", enables the legend, and saves the workbook. | Show how to create a pivot chart in Aspose.Cells by calling PivotTable.NSeries.Add, customize its type, title, and legend, then persist the changes to a new file. | Adapt the example to place the pivot chart on a separate worksheet and switch the chart type to a line chart while keeping the data source linked to the original pivot table.
// Common Searches: Aspose.Cells C# add pivot chart to existing worksheet from pivot table | How to link a column chart to a pivot table using Aspose.Cells for .NET | C# example for creating a pivot chart with title and legend in an Excel file | Save workbook with new pivot chart using Aspose.Cells API | Change pivot chart type to line chart on a new sheet Aspose.Cells C#
// Tags: add pivot chart Aspose.Cells | column chart from PivotTable C# | chart title legend Aspose.Cells | save workbook with chart C# | pivot chart on new worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// Loads an existing workbook, retrieves the first pivot table, adds a column chart linked to that pivot table on the same sheet, sets a title and legend, and saves the workbook as a new file.
class PivotChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that already contains a pivot table with aggregated sales data
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first pivot table (adjust index if needed)
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }
            PivotTable pivot = sheet.PivotTables[0];

            // Add a new chart to the same worksheet (position and size can be adjusted)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Aggregated Sales";

            // Use the pivot table as the source for the chart series
            // The second parameter (true) indicates that the series are plotted by categories
            chart.NSeries.Add(pivot.Name, true);

            // Optional: display legend
            chart.ShowLegend = true;

            // Save the workbook with the new pivot chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
