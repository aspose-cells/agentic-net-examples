using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Optional: customize chart appearance
            chart.Title.Text = "Fruit Sales";
            chart.ShowLegend = true;
            chart.Style = 2; // Built‑in style

            // Create ImageOrPrintOptions with custom attributes
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,          // Output format
                HorizontalResolution = 300,         // DPI X
                VerticalResolution = 300,           // DPI Y
                Quality = 90,                       // Quality (ignored for PNG but kept for completeness)
                Transparent = false                 // No transparency
            };

            // Export the chart to an image file using the custom options
            string imagePath = "FruitSalesChart.png";
            chart.ToImage(imagePath, imgOptions);

            // Save the workbook (optional, to keep the source file)
            workbook.Save("FruitSalesWorkbook.xlsx");

            Console.WriteLine($"Chart image saved to: {imagePath}");
        }
    }
}