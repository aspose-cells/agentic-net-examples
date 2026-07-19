// Title: Lock an Excel chart programmatically with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add sample data, insert a column chart, set ChartObject.IsLocked to true, protect the worksheet, and save as LockedChart.xlsx so the chart cannot be moved or resized in the Excel UI.
// Keywords: Aspose.Cells | C# | lock chart | ChartObject.IsLocked | protect worksheet | prevent chart resizing | Excel chart security
// Common Searches: Aspose.Cells lock chart C# | prevent moving chart Excel Aspose | ChartObject.IsLocked example | protect worksheet to disable chart editing | lock chart shape programmatically
// Developer Intent: Prevent users from moving or resizing a chart in Excel.
// Use Cases: Financial reports where chart positions must stay fixed while data is updated. | Template workbooks that lock chart layout but allow cell edits. | Shared dashboards that let end‑users change values but not alter chart placement.
// AI Prompts: Generate C# code to lock all charts on a worksheet using Aspose.Cells. | Show how to protect a worksheet with custom options while keeping charts locked in Aspose.Cells. | Explain how to unlock a chart that was previously locked with ChartObject.IsLocked.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add sample data, insert a column chart, set ChartObject.IsLocked to true, protect the worksheet, and save as LockedChart.xlsx so the chart cannot be moved or resized in the Excel UI.
class LockChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B3", true);
        chart.Title.Text = "Sample Chart";

        // Lock the chart shape so it cannot be moved or resized when the sheet is protected
        chart.ChartObject.IsLocked = true;

        // Protect the worksheet (default settings disallow editing objects)
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedChart.xlsx");
    }
}
