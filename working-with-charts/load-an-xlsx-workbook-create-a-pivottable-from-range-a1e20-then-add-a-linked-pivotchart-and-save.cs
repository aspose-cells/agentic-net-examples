// Title: Aspose.Cells for .NET – Build a PivotTable from A1:E20, attach a PivotChart, and save the XLSX file
// Description: Load an existing workbook, add a PivotTable on range A1:E20 (placed at G3), define row and data fields, refresh the table, create a column PivotChart linked to the table, refresh the chart, and save the workbook as a new XLSX file using Aspose.Cells.
// Keywords: Aspose.Cells PivotTable C# | link PivotChart to PivotTable Aspose | create column chart from pivot Aspose.Cells | save workbook with pivot chart .NET | Aspose.Cells example PivotTable PivotChart
// Common Searches: Aspose.Cells add PivotTable programmatically | How to bind a chart to a PivotTable with Aspose.Cells | C# create PivotChart from PivotTable using Aspose | Refresh PivotChart after PivotTable changes Aspose.Cells | Save XLSX with linked PivotChart .NET
// Developer Intent: Generate an XLSX workbook that contains a PivotTable built from a specific range and a chart automatically linked to that PivotTable.
// Use Cases: Automated sales dashboards where the PivotTable summarizes raw data and the linked chart updates instantly. | Financial reporting templates that combine a PivotTable summary with a visual trend chart for monthly reviews. | Dynamic Excel dashboards for business intelligence that refresh chart visuals whenever the underlying PivotTable recalculates.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotTable from A1:E20, add a column PivotChart linked to it, and save the workbook. | Explain how to set the PivotSource property of a chart to reference a newly created PivotTable in Aspose.Cells. | Provide troubleshooting steps when a PivotChart does not reflect changes made to its source PivotTable using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace PivotChartExample
{
    // Load an existing workbook, add a PivotTable on range A1:E20 (placed at G3), define row and data fields, refresh the table, create a column PivotChart linked to the table, refresh the chart, and save the workbook as a new XLSX file using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or specify by name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a PivotTable based on the range A1:E20, place it starting at cell G3, and name it "MyPivotTable"
            int pivotIndex = worksheet.PivotTables.Add("A1:E20", "G3", "MyPivotTable");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the PivotTable (example: first column as Row field, second column as Data field)
            // Adjust field indices or names according to your source data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // First column as Row
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Second column as Data

            // Refresh and calculate the PivotTable data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a chart (Column chart) to the worksheet
            // Chart will be placed from row 15, column 0 to row 25, column 7 (adjust as needed)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 15, 0, 25, 7);
            Chart chart = worksheet.Charts[chartIndex];

            // Link the chart to the PivotTable by setting the PivotSource property
            // Assuming the worksheet name is "Sheet1"; adjust if different
            string sheetName = worksheet.Name;
            chart.PivotSource = $"[{workbook.FileName}]'{sheetName}'!{pivotTable.Name}";

            // Refresh the chart to reflect the PivotTable data
            chart.RefreshPivotData();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
