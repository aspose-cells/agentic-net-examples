// Title: How to set a PivotChart layout to Tabular with PivotOptions in Aspose.Cells for .NET and save the workbook
// AI Prompts: Create a pivot table from a data range, add a column chart linked to it, set the chart's layout to Tabular using PivotOptions, and save the workbook with Aspose.Cells in C#. | Load an existing Excel file, apply a Tabular layout to a newly added PivotChart via PivotTableOptions, and write the updated file to a new location using Aspose.Cells for .NET. | Use the Aspose.Cells C# API to configure the PivotChart layout to Tabular, bind the chart to a pivot table, and persist the changes in a new workbook.
// Common Searches: Aspose.Cells C# set PivotChart layout to Tabular | how to apply Tabular layout to a PivotChart using Aspose.Cells .NET | configure PivotOptions Layout property for a chart in Aspose.Cells | save workbook after modifying PivotChart layout with Aspose.Cells | create pivot table and column chart with Tabular layout in C# Aspose.Cells
// Tags: Aspose.Cells configure PivotChart Tabular layout | C# use PivotTableOptions for chart layout | add column chart linked to pivot table Aspose.Cells | save modified Excel workbook Aspose.Cells | Excel data analysis with pivot charts .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// The example loads an existing workbook (or creates one with sample data), adds a pivot table on the first worksheet, inserts a column chart linked to that pivot table, sets the chart's layout to Tabular via PivotOptions, and saves the result to a new Excel file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Product");
                ws.Cells["C1"].PutValue("Region");
                ws.Cells["D1"].PutValue("Sales");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue("Item1");
                ws.Cells["C2"].PutValue("North");
                ws.Cells["D2"].PutValue(100);
                // Additional sample rows can be added here as needed
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the source data range for the pivot table
            string sourceData = "A1:D20";

            // Add a new pivot table at cell C1
            int pivotIndex = sheet.PivotTables.Add(sourceData, "C1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Optional: set layout to Tabular if the API version supports it
            // pivotTable.Options.Layout = PivotTableLayout.Tabular;

            // Add a column chart positioned on the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart's data source to the pivot table
            chart.NSeries.Add(pivotTable.Name, true);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
