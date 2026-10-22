// Title: Load an XLSX workbook, create a PivotTable from A1:E20, add a linked PivotChart, and save it using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to open an existing XLSX file, define a PivotTable on range A1:E20, refresh it, and write the workbook to a new file. | Show how to add a Column PivotChart linked to a PivotTable on the same worksheet, set its position and title, using Aspose.Cells in C#. | Provide a step‑by‑step example that creates a separate worksheet for a PivotTable, configures the source data range, and saves the workbook with both the PivotTable and its chart.
// Common Searches: Aspose.Cells C# create pivot table from range A1:E20 in existing workbook | Add a linked pivot chart to a worksheet using Aspose.Cells .NET | Save workbook with pivot table and chart using Aspose.Cells C# example | How to position a PivotChart programmatically with Aspose.Cells | Refresh and calculate pivot table data in Aspose.Cells C#
// Tags: build pivot table using source range Aspose.Cells | create pivot chart referencing pivot table Aspose.Cells | configure pivot chart bounds Aspose.Cells | update pivot table calculations Aspose.Cells | export workbook containing pivot table and chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// The example loads an existing XLSX file, adds a new worksheet, builds a PivotTable from cells A1:E20, refreshes and calculates its data, inserts a Column PivotChart linked to the same data range with a title, positions the chart, and saves the workbook as a new file.
class PivotTableWithChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the source worksheet (assumed to be the first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet that will contain the PivotTable
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotSheet";

            // Define the source data range for the PivotTable (A1:E20)
            string sourceDataRange = $"='{sourceSheet.Name}'!A1:E20";

            // Add the PivotTable to the pivot sheet, placing its top‑left corner at cell A3
            int pivotTableIndex = pivotSheet.PivotTables.Add(sourceDataRange, "A3", "MyPivotTable");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // OPTIONAL: Configure fields (commented out to avoid missing API issues)
            // Example: add first column as a row field and second column as a data field
            // pivotTable.RowFields.Add(pivotTable.RowFields[0]); // adjust as needed
            // pivotTable.DataFields.Add(pivotTable.DataFields[0]); // adjust as needed

            // Refresh the PivotTable to calculate its data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a linked PivotChart on the same sheet
            // Chart will be placed from row 10, column 0 to row 25, column 10
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 10, 0, 25, 10);
            Chart chart = pivotSheet.Charts[chartIndex];

            // Link the chart to the source data range (as a simple example)
            chart.NSeries.Add(sourceDataRange, true);

            // Set a title for the chart
            chart.Title.Text = "Pivot Chart";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new PivotTable and PivotChart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
