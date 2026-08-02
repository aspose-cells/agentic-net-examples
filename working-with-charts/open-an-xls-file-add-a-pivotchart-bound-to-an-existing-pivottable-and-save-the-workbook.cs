// Title: Add a PivotChart to an Existing PivotTable in an XLS Workbook – Aspose.Cells for .NET
// Description: Loads an XLS file, finds the first PivotTable on the first worksheet, creates a column PivotChart, links the chart to the PivotTable via the PivotSource property, refreshes both the table and the chart, and saves the result as an XLSX workbook using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | PivotChart | PivotTable | bind chart to pivot | refresh pivot data | add chart programmatically | XLS to XLSX conversion | column chart | Excel automation
// Common Searches: Aspose.Cells create PivotChart from existing PivotTable | C# bind chart to PivotTable programmatically | refresh PivotTable and chart after adding PivotChart | convert XLS with PivotChart to XLSX using Aspose | how to add a column chart to a PivotTable with Aspose.Cells
// Developer Intent: Programmatically attach a PivotChart to a pre‑existing PivotTable in a loaded XLS workbook and export the updated file.
// Use Cases: Add a visual sales summary to a legacy XLS report before converting it to XLSX. | Automate monthly dashboards by generating column charts from PivotTables in bulk. | Refresh PivotTable calculations after data changes and ensure the chart reflects the latest values.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line PivotChart to the second PivotTable in a workbook and sets a custom chart title. | Show how to bind a Pie chart to a PivotTable and export the workbook as a PDF using Aspose.Cells. | Explain the steps to update a PivotChart's source range after renaming the underlying PivotTable in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotChartDemo
{
    // Loads an XLS file, finds the first PivotTable on the first worksheet, creates a column PivotChart, links the chart to the PivotTable via the PivotSource property, refreshes both the table and the chart, and saves the result as an XLSX workbook using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Load the existing XLS workbook
            Workbook workbook = new Workbook("input.xls");

            // Assume the first worksheet contains the required PivotTable
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable in the worksheet
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a new chart (Column type) to the same worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Bind the chart to the existing PivotTable
            chart.PivotSource = $"{worksheet.Name}!{pivotTable.Name}";

            // Refresh the PivotTable data and the chart
            pivotTable.RefreshData();
            pivotTable.CalculateData();
            chart.RefreshPivotData();

            // Save the workbook with the new PivotChart
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
