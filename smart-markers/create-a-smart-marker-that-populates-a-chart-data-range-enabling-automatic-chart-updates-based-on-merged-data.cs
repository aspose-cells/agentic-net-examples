// Title: Use Aspose.Cells smart markers to populate a table and auto‑expand a column chart’s named range with merged headers in C#
// AI Prompts: Generate C# code that creates a worksheet with smart‑marker placeholders, processes a line‑by‑line data source, and rebuilds the named range to include all rows. | Show how to add a column chart that references a named range and recalculate it after the smart‑marker processing so the chart updates automatically. | Demonstrate handling merged header cells while using smart markers and updating the chart data range in Aspose.Cells for .NET.
// Common Searches: aspnet smart markers line by line populate chart data range | c# Aspose.Cells dynamic named range for column chart after smart marker processing | how to handle merged header cells with Aspose.Cells smart markers | auto update chart when smart markers fill worksheet Aspose.Cells | create column chart linked to named range using Aspose.Cells .NET
// Tags: smart markers line‑by‑line data population | dynamic named range for Aspose chart | column chart linked to named range | merged header cells handling | chart auto‑recalculation after smart marker processing

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, inserts smart‑marker placeholders for Category and Value, merges the header cells, defines a named range, adds a column chart linked to that range, processes a line‑by‑line data source, rebuilds the named range to cover all populated rows, recalculates the chart so it reflects the new data, and saves the file as an Excel workbook.
class SmartMarkerChartDemo
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- Prepare the template with smart markers ----------
        // Header row
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        // Data rows – smart markers will be repeated line‑by‑line
        // &=$Data.Category and &=$Data.Value will be replaced by the data source values
        sheet.Cells["A2"].PutValue("&=$Data.Category");
        sheet.Cells["B2"].PutValue("&=$Data.Value");

        // Merge the header cells to illustrate merged cells handling
        sheet.Cells.Merge(0, 0, 1, 2); // Merges A1:B1

        // ---------- Add a chart that will use the filled data ----------
        // Add a column chart positioned below the data table
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define a named range that will later cover the filled data (including the header)
        // Initially it points to the placeholder range A1:B2
        sheet.Cells.CreateRange("A1:B2").Name = "ChartData";

        // Set the chart to use the named range; the range will be automatically updated after data is populated
        chart.SetChartDataRange("ChartData", true);
        chart.Title.Text = "Smart Marker Chart";

        // ---------- Prepare the data source ----------
        var data = new List<object>
        {
            new { Category = "Alpha",   Value = 120 },
            new { Category = "Beta",    Value = 150 },
            new { Category = "Gamma",   Value = 90  },
            new { Category = "Delta",   Value = 200 }
        };

        // ---------- Process smart markers ----------
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            LineByLine = true   // Enables line‑by‑line repetition of the smart‑marker rows
        };
        designer.SetDataSource("Data", data);
        designer.Process();   // Fills the smart‑marker cells with the data source values

        // ---------- Adjust the named range to cover all populated rows ----------
        // After processing, MaxDataRow points to the last row containing data
        int lastRow = sheet.Cells.MaxDataRow; // zero‑based index
        // Re‑create the named range to span from A1 to the last data row (including the header)
        sheet.Cells.CreateRange(0, 0, lastRow + 1, 2).Name = "ChartData";

        // Recalculate the chart so it picks up the updated data range
        chart.Calculate();

        // ---------- Save the result ----------
        workbook.Save("SmartMarkerChart.xlsx");
    }
}
