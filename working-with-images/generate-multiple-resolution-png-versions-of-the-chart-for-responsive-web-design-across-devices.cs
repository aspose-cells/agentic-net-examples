using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class GenerateChartImages
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Sales Chart";

        // Define the DPI values for which PNG images will be generated
        int[] dpiValues = new int[] { 96, 150, 300 };

        // Generate a PNG image for each DPI setting
        foreach (int dpi in dpiValues)
        {
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = ImageType.Png;
            options.HorizontalResolution = dpi;
            options.VerticalResolution = dpi;

            string fileName = $"Chart_{dpi}dpi.png";
            chart.ToImage(fileName, options);
            Console.WriteLine($"Chart image saved: {fileName} ({dpi} DPI)");
        }

        // Save the workbook (optional, for reference)
        workbook.Save("ChartWorkbook.xlsx");
    }
}