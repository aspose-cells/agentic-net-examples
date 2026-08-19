// Title: Add a PivotChart to an Existing PivotTable in an XLS Workbook and Save as XLSX – C# with Aspose.Cells
// Description: Loads an XLS file, adds a column PivotChart linked to a PivotTable named "PivotTable1", refreshes the chart and pivot data, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# PivotChart | bind chart to pivot table Aspose | refresh pivot data Aspose.Cells | convert XLS to XLSX with chart | .NET add PivotChart to existing PivotTable | Aspose.Cells chart pivot source | C# Excel pivot chart automation
// Common Searches: How to add a PivotChart to an existing PivotTable with Aspose.Cells | Aspose.Cells refresh pivot table after adding chart | Convert legacy XLS to XLSX and keep PivotChart in C# | C# code to bind chart to PivotTable using Aspose.Cells | Create column chart from PivotTable programmatically
// Developer Intent: Create a column PivotChart bound to an existing PivotTable in an XLS workbook and export the result as an XLSX file using Aspose.Cells for .NET.
// Use Cases: Enhance legacy Excel reports by adding a visual PivotChart before distribution. | Automate migration of old XLS files to XLSX while inserting a summary chart for quick insights. | Refresh pivot data after source updates and generate a ready‑to‑share XLSX workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line PivotChart to a PivotTable called "SalesPivot" in an XLS file and saves it as XLSX. | Explain the steps to bind a chart to a PivotTable and refresh its data using Aspose.Cells for .NET, including required properties and methods. | Provide a step‑by‑step tutorial for loading an XLS workbook, inserting a PivotChart, refreshing pivot data, and converting the file to XLSX with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// Loads an XLS file, adds a column PivotChart linked to a PivotTable named "PivotTable1", refreshes the chart and pivot data, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Name of the existing pivot table to bind the chart to
            string pivotTableName = "PivotTable1";

            // Add a new column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Bind the chart to the existing pivot table
            chart.PivotSource = $"{worksheet.Name}!{pivotTableName}";

            // Refresh the chart data from the pivot table
            chart.RefreshPivotData();

            // Refresh the pivot table data itself, if the pivot table exists
            PivotTable pivotTable = worksheet.PivotTables[pivotTableName];
            if (pivotTable != null)
            {
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }
            else
            {
                Console.WriteLine($"Pivot table '{pivotTableName}' not found in worksheet '{worksheet.Name}'.");
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
