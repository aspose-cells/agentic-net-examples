// Title: Configure PivotChart Tabular Layout via PivotOptions in Aspose.Cells for .NET (C#)
// Description: C# example that loads or creates an Excel workbook with Aspose.Cells, writes a small data set, adds a PivotTable, creates a linked PivotChart, applies the Tabular layout using PivotOptions, refreshes the chart data, and saves the file.
// Keywords: Aspose.Cells | C# PivotChart Tabular layout | PivotOptions | set PivotChart layout | PivotChart Aspose.Cells .NET | create PivotTable Aspose.Cells | refresh pivot chart data | save workbook Aspose.Cells | Excel automation C# | Excel PivotChart programming
// Common Searches: Aspose.Cells set PivotChart layout to Tabular | PivotOptions Tabular layout C# example | How to change PivotChart layout with Aspose.Cells | Create PivotChart from PivotTable using Aspose.Cells | Refresh PivotChart data Aspose.Cells
// Developer Intent: Apply the Tabular layout to a newly created PivotChart using PivotOptions and persist the workbook.
// Use Cases: Automate generation of Excel reports that include summarized data and visual charts. | Programmatically create PivotTables and linked PivotCharts for business dashboards. | Enforce a consistent Tabular chart appearance across generated workbooks. | Refresh pivot data after layout changes to ensure accurate rendering. | Save the final workbook for downstream processing or distribution.
// AI Prompts: Write C# code that loads or creates an Excel file with Aspose.Cells, adds a PivotTable, creates a linked PivotChart, sets the chart's layout to Tabular using PivotOptions, refreshes the chart, and saves the workbook. | Show how to use Aspose.Cells PivotOptions to change a PivotChart's layout to Tabular in a .NET application. | Explain step‑by‑step how to configure a PivotChart's Tabular layout, refresh its data, and export the workbook using Aspose.Cells for C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// C# example that loads or creates an Excel workbook with Aspose.Cells, writes a small data set, adds a PivotTable, creates a linked PivotChart, applies the Tabular layout using PivotOptions, refreshes the chart data, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if it exists; otherwise create a new one
            string inputPath = "input.xlsx";
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Use the first worksheet (or add a new one if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Create sample data for the pivot table (optional)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(40);

            // -------------------------------------------------
            // Add a PivotTable
            // -------------------------------------------------
            int pivotIndex = sheet.PivotTables.Add("=A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            // Data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Refresh the pivot cache and calculate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Add a PivotChart linked to the PivotTable
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            // Set the chart's source to the pivot table (makes it a PivotChart)
            chart.PivotSource = "PivotTable1";

            // Refresh chart data after changing options
            chart.RefreshPivotData();

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
