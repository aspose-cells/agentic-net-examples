// Title: C# – Load XLSX, Create PivotTable (A1:E20), Add Linked Column PivotChart, Save – Aspose.Cells
// Description: Loads an existing XLSX workbook, inserts a PivotTable from range A1:E20 at cell G3, configures row and data fields, refreshes the table, creates a column chart linked to the PivotTable, updates the chart data, and saves the file as output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PivotTable C# | link PivotChart to PivotTable Aspose | create column chart from PivotTable | load and save XLSX Aspose.Cells | refresh PivotTable data Aspose | .NET Excel automation | programmatic Excel dashboard
// Common Searches: How to add a PivotTable from a range with Aspose.Cells .NET | Aspose.Cells example linking a chart to a PivotTable | Create and refresh a column PivotChart in C# using Aspose | Save workbook after adding PivotTable and PivotChart Aspose.Cells
// Developer Intent: Programmatically generate a PivotTable, attach a column PivotChart to it, and persist the workbook.
// Use Cases: Automated sales reporting: generate a summary PivotTable and a column chart for each period. | Dynamic business dashboards: link charts to PivotTables to reflect real‑time data changes. | Batch processing of Excel files: add analytical tables and visualizations before distribution.
// AI Prompts: Write C# code with Aspose.Cells that creates a PivotTable from B2:D30, adds multiple row and data fields, links a line chart to the table, and saves the workbook. | Explain how to change the PivotSource of an existing PivotChart after modifying the PivotTable layout using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Loads an existing XLSX workbook, inserts a PivotTable from range A1:E20 at cell G3, configures row and data fields, refreshes the table, creates a column chart linked to the PivotTable, updates the chart data, and saves the file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Add a PivotTable using the range A1:E20, place it starting at cell G3, and name it "PivotTable1"
        int pivotIndex = sheet.PivotTables.Add("A1:E20", "G3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Example configuration: first column as Row field, second column as Data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        // Refresh and calculate the PivotTable data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Add a Column chart to the worksheet (position can be adjusted)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Link the chart to the created PivotTable
        chart.PivotSource = $"{sheet.Name}!{pivotTable.Name}";

        // Refresh the chart to reflect the PivotTable data
        chart.RefreshPivotData();

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
