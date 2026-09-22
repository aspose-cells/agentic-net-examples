// Title: Generate a PivotTable from A1:E20 and attach a linked Column PivotChart in an XLSX file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing XLSX workbook, creates a PivotTable on a new sheet from the range A1:E20, adds a Column PivotChart linked to that table, and saves the file. | Show how to programmatically link a PivotChart to a PivotTable in Aspose.Cells for .NET, including setting the chart type and title. | Provide a step‑by‑step example of adding a worksheet, inserting a PivotTable, and creating a linked PivotChart using the Aspose.Cells C# API.
// Common Searches: asp.net create pivot table from range A1:E20 using Aspose.Cells | how to add a linked pivot chart to an Excel workbook with Aspose.Cells C# | sample code for generating a column pivot chart from a pivot table in Aspose.Cells | save workbook with pivot table and chart Aspose.Cells .NET example
// Tags: Aspose.Cells pivot table from A1:E20 | Aspose.Cells linked pivot chart | C# column pivot chart Aspose.Cells | save XLSX with pivot table and chart Aspose.Cells | add worksheet and pivot objects Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// // This program loads an existing XLSX file, adds a new worksheet, creates a PivotTable from the source range A1:E20 on the first sheet, inserts a linked Column PivotChart, and saves the modified workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet (assumed first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet for the PivotTable and PivotChart
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotSheet";

            // Define source data range for the PivotTable
            string sourceData = $"{sourceSheet.Name}!$A$1:$E$20";

            // Add the PivotTable to the new sheet at cell A1
            int pivotTableIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // NOTE: Field configuration (RowFields, ColumnFields, DataFields) is omitted
            // to maintain compatibility across different Aspose.Cells versions.

            // Add a PivotChart linked to the PivotTable
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
            Chart chart = pivotSheet.Charts[chartIndex];
            chart.Title.Text = "Pivot Chart";

            // Link the chart to the source data range
            chart.NSeries.Add(sourceData, true);

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
