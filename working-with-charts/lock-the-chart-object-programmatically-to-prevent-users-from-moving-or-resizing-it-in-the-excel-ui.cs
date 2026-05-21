using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLockDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);
            chart.Title.Text = "Sample Chart";

            // Obtain the ChartShape (the underlying shape of the chart)
            ChartShape chartShape = chart.ChartObject;

            // Lock the shape so it cannot be moved or resized when the sheet is protected
            chartShape.IsLocked = true; // General lock
            chartShape.SetLockedProperty(ShapeLockType.Move, true);    // Prevent moving
            chartShape.SetLockedProperty(ShapeLockType.Resize, true);  // Prevent resizing

            // Protect the worksheet (required for the lock to take effect)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("ChartLockedDemo.xlsx");
        }
    }
}