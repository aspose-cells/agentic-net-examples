// Title: Create a PivotChart with Multiple Value Fields and Custom Series Colors in C# using Aspose.Cells
// Description: Loads an existing XLS workbook, adds a pivot table with one row field and two data fields, generates a column PivotChart, assigns distinct colors to each series, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | PivotChart | multiple value fields | custom series colors | load XLS workbook | save XLSX | column chart | pivot table chart series color
// Common Searches: Aspose.Cells create pivot chart with multiple data fields | set different colors for pivot chart series in C# | load XLS and add pivot table then chart Aspose.Cells | refresh pivot chart after modifying pivot table .NET | example of custom series colors in Aspose.Cells pivot chart
// Developer Intent: Add a column PivotChart that reflects a pivot table with two value fields and apply individual colors to each series, then save the workbook.
// Use Cases: Build a sales dashboard where regions are rows and revenue & units sold appear as blue and orange columns. | Automate legacy XLS financial reports by inserting a pivot table with expense categories and two metrics, then export a colored chart to XLSX. | Create a data‑analysis utility that reads raw data, generates a multi‑field pivot table, visualizes it with a colored PivotChart, and delivers the file to end users.
// AI Prompts: Generate C# code with Aspose.Cells that loads an XLS file, creates a pivot table with three data fields, builds a line PivotChart, and sets each series to a specific RGB color. | Explain how to refresh a PivotChart after updating its source pivot table when using Aspose.Cells for .NET. | Provide a robust example that validates the input file, creates a pivot chart with custom series colors, and includes comprehensive exception handling.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Loads an existing XLS workbook, adds a pivot table with one row field and two data fields, generates a column PivotChart, assigns distinct colors to each series, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputData.xls";
            const string outputPath = "OutputWithPivotChart.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet (assumed to contain the source data)
            Worksheet sheet = workbook.Worksheets[0];

            // Add a pivot table based on a data range (adjust the range as needed)
            int pivotIdx = sheet.PivotTables.Add("A1:C10", "E3", "MyPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Configure the pivot fields:
            // - Row area uses the first column (index 0)
            // - Data area uses the second and third columns (indexes 1 and 2) -> multiple value fields
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            pivot.AddFieldToArea(PivotFieldType.Data, 1);
            pivot.AddFieldToArea(PivotFieldType.Data, 2);

            // Populate the pivot table with calculated data
            pivot.CalculateData();

            // Add a column chart that will be linked to the pivot table (pivot chart)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the pivot source of the chart to the newly created pivot table
            chart.PivotSource = "MyPivot";

            // Refresh the chart so it reflects the pivot data and creates series
            chart.RefreshPivotData();

            // Assign individual colors to each series (each value field becomes a series)
            if (chart.NSeries.Count >= 2)
            {
                chart.NSeries[0].Area.ForegroundColor = Color.Blue;
                chart.NSeries[1].Area.ForegroundColor = Color.Orange;
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
