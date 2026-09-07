// Title: How to lock a chart shape in Aspose.Cells C# so it cannot be moved or resized in Excel
// AI Prompts: Write C# code with Aspose.Cells that adds a chart, sets ChartShape.IsLocked to true, applies ShapeLockType for Move and Resize, and then protects the worksheet. | Show a step‑by‑step example of preventing users from moving or resizing a chart in an Excel file by configuring chart shape locking and worksheet protection using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# lock chart position and size in Excel workbook | prevent users from moving a chart in Excel using Aspose.Cells API | set chart shape IsLocked and ShapeLockType in Aspose.Cells .NET | how to protect worksheet to enforce chart lock with Aspose.Cells | C# example for disabling chart resizing in generated Excel file
// Tags: chart shape lock Aspose.Cells C# | disable chart move resize Aspose.Cells | worksheet protection chart lock Aspose.Cells | set ShapeLockType Move Resize Aspose.Cells | lock chart object Excel Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLockDemo
{
    // The sample creates a workbook, adds sample data, inserts a column chart, obtains its ChartShape, sets IsLocked to true, locks move and resize via ShapeLockType, protects the worksheet, and saves the file as LockedChartDemo.xlsx.
    class Program
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
