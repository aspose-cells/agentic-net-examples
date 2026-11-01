// Title: Create a PivotChart linked to an existing PivotTable in an XLS workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xls file with Aspose.Cells, locates the first PivotTable, adds a column PivotChart bound to that PivotTable, sets a chart title, and saves the workbook. | Show how to use Aspose.Cells Chart.NSeries.Add method to bind a chart series to a PivotTable name and configure the chart type in C#. | Demonstrate error handling in C# for scenarios where the input .xls file is missing or the worksheet contains no PivotTables before creating a PivotChart with Aspose.Cells.
// Common Searches: aspnet how to add a pivotchart to an existing xls file using Aspose.Cells | c# bind chart series to a pivot table name with Aspose.Cells | create column chart from pivot table in an xls workbook Aspose.Cells example | asp.net check for pivot tables before adding chart with Aspose.Cells
// Tags: add pivotchart to xls workbook Aspose.Cells | bind chart series to pivot table Aspose.Cells C# | column pivotchart creation Aspose.Cells | load and save xls with chart Aspose.Cells | missing pivot table error handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// The example loads an existing XLS workbook, retrieves the first PivotTable on the first worksheet, adds a column PivotChart bound to that PivotTable, sets a chart title, and saves the modified workbook as a new XLS file, including basic error handling for missing files or PivotTables.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xls";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable on the worksheet
            // (Assumes at least one PivotTable exists)
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found on the first worksheet.");
                return;
            }
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a new chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set a title for the chart
            chart.Title.Text = "Pivot Chart";

            // Bind the chart to the PivotTable
            // The first argument is the name of the PivotTable, the second indicates that the series are in columns
            chart.NSeries.Add(pivotTable.Name, true);

            // Save the workbook with the new PivotChart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
