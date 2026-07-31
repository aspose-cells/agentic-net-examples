// Title: Aspose.Cells for .NET – Create a PivotTable from A1:E20 and link a PivotChart (C#)
// Description: Load an XLSX file, add a PivotTable based on range A1:E20 with the report starting at G3, configure row and data fields, refresh the table, create a linked column PivotChart, refresh the chart, and save the workbook as output.xlsx using Aspose.Cells C# API.
// Keywords: Aspose.Cells C# PivotTable | Aspose.Cells linked PivotChart | create pivot table programmatically | add chart to pivot table .NET | XLSX pivot example | Aspose.Cells chart API | C# Excel automation | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells add PivotTable from range | C# link PivotChart to PivotTable | create pivot chart programmatically Aspose.Cells | Aspose.Cells example for PivotTable and chart | how to refresh pivot chart Aspose.Cells
// Developer Intent: Generate a PivotTable from a defined range, attach a linked PivotChart, and save the updated workbook.
// Use Cases: Automated sales or financial reporting where a pivot summary and chart are generated on the fly. | Server‑side Excel processing to produce dashboards with synchronized pivot tables and charts. | Batch processing of uploaded spreadsheets to add standardized pivot analyses and visualizations.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotTable from A1:E20, set row and data fields, and add a linked column chart. | Explain how to refresh a PivotChart after updating its source PivotTable using Aspose.Cells. | Show how to customize the style and layout of a PivotChart linked to a PivotTable in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Load an XLSX file, add a PivotTable based on range A1:E20 with the report starting at G3, configure row and data fields, refresh the table, create a linked column PivotChart, refresh the chart, and save the workbook as output.xlsx using Aspose.Cells C# API.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Add a PivotTable:
        //   Source range: A1:E20
        //   Destination cell (upper‑left corner of the report): G3
        //   Table name: MyPivot
        int pivotIndex = sheet.PivotTables.Add("A1:E20", "G3", "MyPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the PivotTable (example: first column as Row field, second column as Data field)
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column A
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column B

        // Refresh and calculate the PivotTable data
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a chart (Column type) and link it to the PivotTable
        int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 25, 7);
        Chart chart = sheet.Charts[chartIndex];
        chart.PivotSource = $"{sheet.Name}!{pivot.Name}";

        // Refresh the chart to reflect the PivotTable data
        chart.RefreshPivotData();

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
