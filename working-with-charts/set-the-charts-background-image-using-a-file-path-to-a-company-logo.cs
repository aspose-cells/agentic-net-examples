// Title: Set Chart Background Image from a File Path (Company Logo) with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds sample data, inserts a column chart, and applies a PNG logo as the chart area's background using texture fill. The code checks the file path, loads the image into a byte array, sets FillFormat.TextureFill.ImageData, and saves the workbook as an .xlsx file.
// Keywords: Aspose.Cells chart background image | C# set chart texture fill | load PNG into chart area Aspose | Excel chart logo background .NET | FillFormat.TextureFill ImageData | Aspose.Cells file path image
// Common Searches: Aspose.Cells set chart background from file | C# add logo to Excel chart background | texture fill chart area Aspose.Cells | how to use FillFormat.TextureFill in .NET | chart background image not showing Aspose
// Developer Intent: Apply a company logo stored on disk as the background image of an Excel chart using Aspose.Cells for .NET.
// Use Cases: Generate a column chart and brand it with a corporate logo as the background. | Validate the logo file exists before applying the texture fill to avoid runtime errors. | Produce Excel reports where charts carry consistent visual identity across automated workflows.
// AI Prompts: Write C# code that loads a PNG file and sets it as the background image of an Aspose.Cells chart, with error handling for missing files. | Explain the role of FillFormat.TextureFill.ImageData when applying a texture fill to a chart area in Aspose.Cells. | Provide step‑by‑step instructions to embed a company logo as a chart background and export the workbook to .xlsx.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBackgroundImage
{
    // Creates a new workbook, adds sample data, inserts a column chart, and applies a PNG logo as the chart area's background using texture fill. The code checks the file path, loads the image into a byte array, sets FillFormat.TextureFill.ImageData, and saves the workbook as an .xlsx file.
    public class SetChartBackgroundImage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // (Optional) Add some sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["B4"].PutValue(20);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Load the company logo image into a byte array if the file exists
                string logoPath = "company_logo.png"; // Replace with the actual file path
                if (File.Exists(logoPath))
                {
                    byte[] logoBytes = File.ReadAllBytes(logoPath);
                    // Apply texture fill to the chart area and set the image data
                    chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
                    chart.ChartArea.Area.FillFormat.TextureFill.ImageData = logoBytes;
                }
                else
                {
                    Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Skipping background image.");
                }

                // Save the workbook with the chart background image applied
                string outputPath = "ChartWithBackgroundImage.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetChartBackgroundImage.Run();
        }
    }
}
