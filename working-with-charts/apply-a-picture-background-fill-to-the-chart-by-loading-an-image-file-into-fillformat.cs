// Title: Apply a picture background to a chart area using Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet set image as chart background using Aspose.Cells FillFormat | how to use texture fill to add a PNG background to an Excel chart with Aspose.Cells | Aspose.Cells chart area picture fill example in C# | load image bytes into FillFormat.TextureFill for Excel chart background .NET | set FillPictureType.Stretch for chart background image Aspose.Cells
// Tags: chart area texture fill Aspose.Cells | image background fill for Excel chart C# | FillFormat.TextureFill image data Aspose.Cells | FillPictureType.Stretch chart background .NET | apply picture fill to chart area Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data and a column chart, reads a PNG file into a byte array, sets the chart area's FillFormat to Texture, assigns the image bytes to the texture fill, configures the picture format to stretch, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Path to the picture that will be used as the chart background.
        const string imagePath = "background.png";

        // Verify that the image file exists.
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Error: Image file \"{imagePath}\" not found.");
            return;
        }

        try
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and obtain the first worksheet.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 2. Populate some sample data for the chart.
            // ------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // ------------------------------------------------------------
            // 3. Add a column chart to the worksheet.
            // ------------------------------------------------------------
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart.
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ------------------------------------------------------------
            // 4. Apply a picture background fill to the entire chart.
            // ------------------------------------------------------------
            // Set the fill type to Texture (used for picture fills).
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load the image bytes safely.
            byte[] imageBytes;
            try
            {
                imageBytes = File.ReadAllBytes(imagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read image file: {ex.Message}");
                return;
            }

            // Assign the image data to the texture fill.
            chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageBytes;

            // Define how the picture is rendered (stretch to fill the chart area).
            chart.ChartArea.Area.FillFormat.TextureFill.PictureFormatType = FillPictureType.Stretch;
            chart.ChartArea.Area.FillFormat.Scale = 1.0; // 100% scale

            // ------------------------------------------------------------
            // 5. Save the workbook.
            // ------------------------------------------------------------
            const string outputFile = "ChartWithPictureBackground.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
