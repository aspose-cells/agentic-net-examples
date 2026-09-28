// Title: How to set a chart's background image from a file path using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart in an Aspose.Cells workbook and configures the chart area's FillFormat to use a texture fill. | Write C# that reads a PNG logo file from a specified path into a byte array and assigns it to chart.ChartArea.Area.FillFormat.TextureFill.ImageData.
// Common Searches: Aspose.Cells C# set chart background image from local file | apply texture fill to chart area using Aspose.Cells .NET | load PNG file into byte array for chart fill Aspose.Cells | add company logo as background to Excel chart with Aspose.Cells
// Tags: chart area texture fill Aspose.Cells | set chart background image C# | load image bytes for chart fill Aspose.Cells | apply PNG logo to chart background | column chart background image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBackgroundImage
{
    // The example creates a new workbook, adds sample data, inserts a column chart, changes the chart area's fill type to texture, reads a PNG logo from disk into a byte array, assigns the image as the chart's background, and saves the workbook as ChartWithBackgroundImage.xlsx.
    public class SetChartBackgroundImage
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Configure the chart area to use a texture fill (image background)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load the company logo image into a byte array if the file exists
            string logoPath = "company_logo.png"; // replace with actual file path
            if (File.Exists(logoPath))
            {
                try
                {
                    byte[] logoBytes = File.ReadAllBytes(logoPath);
                    chart.ChartArea.Area.FillFormat.TextureFill.ImageData = logoBytes;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to read logo file. {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Skipping background image.");
            }

            // Save the workbook with the chart background image applied
            string outputPath = "ChartWithBackgroundImage.xlsx";
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
                throw;
            }
        }
    }
}
