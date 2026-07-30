// Title: Create a PivotChart with Multiple Value Fields and Custom Series Colors in C# using Aspose.Cells
// Description: Loads an XLS workbook, builds a PivotTable with two data fields, adds a column PivotChart linked to the table, assigns red and blue colors to the first two series, refreshes the chart data, and saves the file as XLSX.
// Keywords: Aspose.Cells PivotChart C# | custom series colors Aspose.Cells | multiple value fields PivotTable | convert XLS to XLSX with chart | chart series foreground color .NET | programmatic PivotChart creation | Aspose.Cells chart customization
// Common Searches: how to set series colors in a PivotChart using Aspose.Cells | create PivotTable with two data fields C# Aspose.Cells | add a colored PivotChart to an existing XLS workbook | convert legacy XLS to XLSX and keep charts | Aspose.Cells change chart series foreground color
// Developer Intent: Generate a PivotChart from an existing XLS file, include two value fields, apply distinct colors to each series, and output the workbook as XLSX.
// Use Cases: Build a sales dashboard where Sales and Quantity appear as separate colored columns in a PivotChart. | Automate migration of old XLS reports to modern XLSX files while adding visually distinct PivotCharts for executive review. | Programmatically enrich a legacy workbook with a PivotChart that uses custom series colors for clearer data presentation.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotTable with three data fields and assign each series a specific RGB color. | Explain how to refresh a PivotChart after updating its source data using Aspose.Cells for .NET. | Show how to change the chart type of a PivotChart and customize markers and series colors in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Loads an XLS workbook, builds a PivotTable with two data fields, adds a column PivotChart linked to the table, assigns red and blue colors to the first two series, refreshes the chart data, and saves the file as XLSX.
class PivotChartWithColors
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Use the first worksheet (adjust if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Create a PivotTable with multiple value fields
        // -------------------------------------------------
        // Define the source data range (adjust to match your file)
        string sourceRange = "A1:E20";

        // Add the PivotTable at cell E3 with the name "PivotTable1"
        int pivotIndex = sheet.PivotTables.Add(sourceRange, "E3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields:
        //   Row field – first column (e.g., Category)
        //   Data fields – third and fourth columns (e.g., Sales and Quantity)
        pivot.AddFieldToArea(PivotFieldType.Row, 0);          // Column A as Row
        pivot.AddFieldToArea(PivotFieldType.Data, 2);         // Column C as first Value
        pivot.AddFieldToArea(PivotFieldType.Data, 3);         // Column D as second Value

        // Populate the PivotTable
        pivot.CalculateData();

        // -------------------------------------------------
        // Create a PivotChart based on the PivotTable
        // -------------------------------------------------
        // Add a column chart (you can choose any ChartType)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Link the chart to the PivotTable – this makes it a PivotChart
        chart.PivotSource = "PivotTable1";

        // Refresh chart data from the PivotTable
        chart.RefreshPivotData();

        // -------------------------------------------------
        // Assign individual colors to each data series
        // -------------------------------------------------
        // Ensure the series exist (one series per value field)
        if (chart.NSeries.Count > 0)
        {
            // First value field – red
            chart.NSeries[0].Area.ForegroundColor = Color.Red;

            // Second value field – blue (if present)
            if (chart.NSeries.Count > 1)
                chart.NSeries[1].Area.ForegroundColor = Color.Blue;
        }

        // -------------------------------------------------
        // Save the modified workbook
        // -------------------------------------------------
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
