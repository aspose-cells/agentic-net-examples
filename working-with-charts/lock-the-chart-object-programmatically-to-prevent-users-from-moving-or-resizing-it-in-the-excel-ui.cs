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

            // Populate some sample data for the chart
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

            // Lock the chart shape so it cannot be moved or resized when the sheet is protected
            // ChartObject returns the underlying ChartShape (inherits from Shape)
            chart.ChartObject.IsLocked = true;

            // Optionally, lock specific properties such as Move and Resize using SetLockedProperty
            // This ensures the chart cannot be moved or resized even if IsLocked is somehow bypassed
            chart.ChartObject.SetLockedProperty(ShapeLockType.Move, true);
            chart.ChartObject.SetLockedProperty(ShapeLockType.Resize, true);

            // Protect the worksheet (all protection options enabled by default)
            // The locked state of the chart takes effect only when the worksheet is protected
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("ChartLockedDemo.xlsx");
        }
    }
}