// Title: Create a PivotChart with custom axis titles and visible drop zones using Aspose.Cells for .NET
// Description: Loads or creates an XLSX file, adds a pivot table, generates a column PivotChart linked to the table, assigns custom text to the category and value axes, enables drop zones via PivotOptions, refreshes the chart data, and saves the workbook.
// Keywords: Aspose.Cells | C# | PivotChart | custom axis title | DropZonesVisible | PivotOptions | Excel automation | .NET chart from pivot table | set chart axis text
// Common Searches: Aspose.Cells set pivot chart axis title C# | Enable drop zones on PivotChart Aspose.Cells | Create PivotChart from pivot table .NET | How to change category and value axis labels in Aspose.Cells | PivotOptions DropZonesVisible example
// Developer Intent: Add a PivotChart, customize its axes, turn on drop zones, and save the workbook programmatically.
// Use Cases: Produce a sales dashboard where the chart axes convey clear metric names. | Automate a reporting workbook that lets end users rearrange fields directly on the chart. | Update legacy Excel files to replace generic axis labels with business‑specific titles while exposing interactive drop zones.
// AI Prompts: Generate C# code with Aspose.Cells that creates a pivot table, links a column chart, sets custom category and value axis titles, enables drop zones, and saves the file. | Explain how to use PivotOptions.DropZonesVisible to make pivot chart drop zones visible in Aspose.Cells for .NET. | Provide a step‑by‑step example of customizing axis titles on a PivotChart and refreshing its data using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// Loads or creates an XLSX file, adds a pivot table, generates a column PivotChart linked to the table, assigns custom text to the category and value axes, enables drop zones via PivotOptions, refreshes the chart data, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Ensure the input workbook exists; create a simple one if missing.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                // Sample data for the pivot source (A1:B5)
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("A");
                ws.Cells["B4"].PutValue(30);
                ws.Cells["A5"].PutValue("B");
                ws.Cells["B5"].PutValue(40);
                wb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a pivot table (source data: A1:B5) at D1.
            int pivotIndex = worksheet.PivotTables.Add("=A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value field
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a column chart linked to the pivot table.
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.PivotSource = "PivotTable1";

            // Set custom axis titles.
            chart.CategoryAxis.Title.Text = "Custom Category Axis";
            chart.ValueAxis.Title.Text = "Custom Value Axis";

            // Configure PivotOptions (make drop zones visible).
            PivotOptions pivotOptions = chart.PivotOptions;
            pivotOptions.DropZonesVisible = true;
            // Note: ShowExpandCollapseButtons property is not available in this version of Aspose.Cells.

            // Refresh chart data from the pivot table.
            chart.RefreshPivotData();

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
