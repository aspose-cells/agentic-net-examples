using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B4", true);

            // Export the chart to a PNG image using default settings
            // The overload with (string, ImageType) saves the image directly to a file
            chart.ToImage("PieChart.png", ImageType.Png);

            // Optionally, save the workbook if you want to keep the Excel file as well
            workbook.Save("PieChartWorkbook.xlsx");

            Console.WriteLine("Pie chart exported to PieChart.png successfully.");
        }
    }
}