// Title: Automatically refresh a pivot‑based chart when an Excel workbook is opened using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program with Aspose.Cells that builds a pivot table, links a column chart to it, and configures the workbook so the chart updates each time the file is opened. | Modify an existing Aspose.Cells workbook to enable RefreshDataOnOpeningFile for its pivot table and refresh the associated chart before saving. | Generate an .xlsx file in C# where the chart source is a pivot table and the chart data is automatically refreshed on workbook open.
// Common Searches: Aspose.Cells C# set chart to refresh automatically on workbook open | how to enable RefreshDataOnOpeningFile for pivot chart in Aspose.Cells | auto update pivot chart when opening Excel file using Aspose.Cells .NET | C# example of chart linked to pivot table with auto refresh on open | Aspose.Cells refresh chart data on file open property
// Tags: auto refresh chart on workbook open Aspose.Cells | pivot table RefreshDataOnOpeningFile C# | chart linked to pivot table Aspose.Cells | generate Excel file with auto updating chart C# | set chart refresh property Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds sample data, builds a pivot table, enables RefreshDataOnOpeningFile so the pivot updates when the file is opened, adds a column chart linked to the pivot table, refreshes the chart data, and saves the workbook as ChartAutoRefreshOnOpen.xlsx.
    public class ChartAutoRefreshOnOpenDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the source data
            int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

            // Ensure the pivot table refreshes its data when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Add a chart that uses the pivot table as its source
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = dataSheet.Charts[chartIndex];
            chart.PivotSource = "Pivot!PivotTable1";

            // Refresh the chart data now so the file contains up‑to‑date data
            chart.RefreshPivotData();

            // Save the workbook; when opened in Excel the chart will reflect the latest data
            workbook.Save("ChartAutoRefreshOnOpen.xlsx", SaveFormat.Xlsx);
        }
    }
}
