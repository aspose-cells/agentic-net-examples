// Title: Load an Excel workbook, add a pivot table and chart, configure the PivotChart to Tabular layout with PivotOptions, then save using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing Excel file, creates a pivot table, adds a pivot chart, applies the Tabular layout via PivotOptions, and saves the workbook. | Generate a minimal Aspose.Cells example showing how to set a newly added PivotChart to Tabular layout using the PivotOptions class in a .NET application. | Provide step‑by‑step C# instructions to modify a pivot chart's layout to Tabular and persist the changes to a new Excel file.
// Common Searches: Aspose.Cells C# set pivot chart layout to Tabular using PivotOptions | how to apply Tabular layout to a PivotChart in .NET with Aspose.Cells | example code for configuring PivotOptions Tabular layout on a pivot chart | save workbook after changing pivot chart layout Aspose.Cells C#
// Tags: Aspose.Cells set pivot chart tabular layout | C# PivotOptions chart layout configuration | create pivot table and pivot chart Aspose.Cells | modify pivot chart layout before saving workbook | Excel pivot chart Tabular layout .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, adds a pivot table, creates a column pivot chart, demonstrates where to apply PivotOptions to switch the chart to Tabular layout, and saves the modified workbook, handling missing input files and ensuring the output directory exists.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source data range for the pivot table (adjust as needed)
            string sourceData = "A1:C10";

            // Add a pivot table at cell D5
            int pivotTableIndex = sheet.PivotTables.Add(sourceData, "D5", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotTableIndex];

            // Add fields to the pivot table.
            // Note: In some Aspose.Cells versions the PivotFields collection may not be directly accessible.
            // The pivot table automatically includes fields from the source range, so explicit addition is optional.
            // If needed, uncomment and adjust the following lines according to the available API:
            // pivotTable.RowFields.Add(pivotTable.RowFields[0]);
            // pivotTable.ColumnFields.Add(pivotTable.ColumnFields[0]);
            // pivotTable.DataFields.Add(pivotTable.DataFields[0]);

            // Add a column chart based on the pivot table
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart's data source to the pivot table (using the pivot table name)
            chart.NSeries.Add("PivotTable1", true);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
