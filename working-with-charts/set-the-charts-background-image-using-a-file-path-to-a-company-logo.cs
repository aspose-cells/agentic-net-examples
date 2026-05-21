using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class SetChartBackgroundImage
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Load the company logo image if it exists
            string logoPath = "company_logo.png";
            if (File.Exists(logoPath))
            {
                byte[] logoBytes = File.ReadAllBytes(logoPath);
                chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = logoBytes;
                chart.ChartArea.BackgroundMode = BackgroundMode.Opaque;
            }
            else
            {
                Console.WriteLine($"Logo file not found: {logoPath}. Skipping background image.");
            }

            // Save the workbook with the chart
            string outputPath = "ChartWithLogoBackground.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}