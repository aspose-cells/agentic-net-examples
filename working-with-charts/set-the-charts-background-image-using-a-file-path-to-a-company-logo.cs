// Title: How to set an Excel chart background image from a file path using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart background image from local file | apply logo as chart area texture fill using Aspose.Cells .NET | load image bytes and use as Excel chart background with Aspose.Cells | C# Aspose.Cells chart background image scaling example | set chart background image path Aspose.Cells workbook
// Tags: chart area texture fill with image Aspose.Cells | apply logo to Excel chart background C# | load PNG bytes for chart fill Aspose.Cells | scale chart background image Aspose.Cells | add column chart and set background image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartBackgroundImageDemo
{
    // The example reads a PNG logo file into a byte array, creates a workbook with sample data, adds a column chart, and sets the chart's background by configuring the ChartArea fill type to Texture and assigning the image bytes. It also demonstrates scaling the background image before saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the logo image that will be used as the chart background.
            // Update this path to point to an existing image file on your machine.
            string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CompanyLogo.png");

            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Error: Logo image not found at \"{logoPath}\".");
                return;
            }

            try
            {
                // Load the logo image bytes.
                byte[] logoBytes = File.ReadAllBytes(logoPath);

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data for the chart.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart.
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Set the chart's background image using texture fill.
                chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = logoBytes;

                // Optional: make the background image stretch to fill the chart area.
                chart.ChartArea.Area.FillFormat.TextureFill.Scale = 1.0; // 100% scale

                // Save the workbook.
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChartWithLogoBackground.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
                Console.WriteLine("The chart's background image has been set to the specified logo.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart with a background image:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
