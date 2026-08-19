// Title: Lock an Excel chart (prevent move & resize) using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds a column chart, retrieves its ChartShape, sets IsLocked and ShapeLockType.Move/Resize to true, then protects the worksheet so the chart cannot be moved or resized in Excel. Saves as LockedChartDemo.xlsx.
// Keywords: Aspose.Cells chart lock | C# lock Excel chart | prevent chart move | disable chart resize | worksheet protection Aspose | ChartShape IsLocked | ShapeLockType Move Resize | programmatic chart security | Excel UI chart lock
// Common Searches: how to lock a chart in Excel with Aspose.Cells C# | prevent chart resizing using Aspose.Cells for .NET | Aspose.Cells protect worksheet chart shape | disable moving of Excel chart programmatically | C# lock chart object Aspose.Cells
// Developer Intent: Secure a chart so end‑users cannot move or resize it in the Excel interface.
// Use Cases: Fixed‑position charts in financial reports shared across teams. | Template workbooks where chart layout must stay consistent. | Collaborative dashboards that require chart positions to remain unchanged.
// AI Prompts: Generate C# code that locks multiple chart objects and applies different ShapeLockType settings with Aspose.Cells. | Explain how worksheet protection interacts with ChartShape.IsLocked and ShapeLockType in Aspose.Cells for .NET. | Show an example of unlocking a chart after protection is removed using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLockDemo
{
    // C# example that creates a workbook, adds a column chart, retrieves its ChartShape, sets IsLocked and ShapeLockType.Move/Resize to true, then protects the worksheet so the chart cannot be moved or resized in Excel. Saves as LockedChartDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);
            chart.Title.Text = "Sample Chart";

            // Obtain the chart's shape (ChartShape) and lock it
            ChartShape chartShape = chart.ChartObject;
            chartShape.IsLocked = true; // Prevents any modification when the sheet is protected
            chartShape.SetLockedProperty(ShapeLockType.Move, true);    // Disallow moving
            chartShape.SetLockedProperty(ShapeLockType.Resize, true);  // Disallow resizing

            // Protect the worksheet so that the lock takes effect
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("LockedChartDemo.xlsx");
        }
    }
}
