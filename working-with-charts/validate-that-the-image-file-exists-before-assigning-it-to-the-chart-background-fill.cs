// Title: Validate image file existence before applying it as a chart background texture in Aspose.Cells for .NET
// Common Searches: c# aspose.cells verify image path before setting chart background texture | aspose.cells chart area texture fill with file existence validation | how to handle missing background image when adding a chart in Aspose.Cells C# | set chart background image from file only if it exists using Aspose.Cells | validate png file before using it as chart background in Aspose.Cells workbook
// Tags: chart area texture fill Aspose.Cells | image file validation Aspose.Cells C# | set chart background image Aspose.Cells | file existence check before chart fill | Aspose.Cells workbook chart background handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace ChartBackgroundImageDemo
{
    // The example checks whether the specified image file exists, reads its bytes, creates a workbook, adds a column chart, and applies the image as a texture fill to the chart area only after successful validation, then saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // Expected usage:
            //   ChartBackgroundImageDemo.exe <imageFilePath> [outputFilePath]
            // If no arguments are supplied, the program uses a default image
            // path "background.png" and saves the workbook as "ChartWithBackground.xlsx".
            // ------------------------------------------------------------

            string imagePath;
            string outputPath;

            if (args.Length >= 1 && !string.IsNullOrWhiteSpace(args[0]))
            {
                imagePath = args[0];
            }
            else
            {
                // Default image file name placed in the same folder as the executable
                imagePath = "background.png";
            }

            if (args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1]))
            {
                outputPath = args[1];
            }
            else
            {
                outputPath = "ChartWithBackground.xlsx";
            }

            // Validate that the image file exists before proceeding
            if (!File.Exists(imagePath))
            {
                Console.Error.WriteLine($"Error: Image file \"{imagePath}\" does not exist.");
                Console.Error.WriteLine("Please provide a valid image file path as the first argument.");
                return;
            }

            try
            {
                // Load image bytes
                byte[] imageBytes = File.ReadAllBytes(imagePath);

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set the chart area fill to use the image as a texture
                // First, ensure the FillFormat object is accessible
                FillFormat chartAreaFill = chart.ChartArea.Area.FillFormat;
                chartAreaFill.FillType = FillType.Texture; // Use texture fill
                chartAreaFill.TextureFill.ImageData = imageBytes; // Assign image data

                // Optionally, adjust transparency or other texture properties
                chartAreaFill.TextureFill.Transparency = 0.0; // Fully opaque

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
                Console.WriteLine($"Chart background image set from \"{imagePath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An unexpected error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
