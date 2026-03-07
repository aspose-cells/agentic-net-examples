using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackgroundDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1) Set a solid background color for the entire sheet
            // -------------------------------------------------
            // Create a style with a solid fill and the desired background color
            Style bgStyle = workbook.CreateStyle();
            bgStyle.Pattern = BackgroundType.Solid;          // Use solid pattern
            bgStyle.BackgroundColor = Color.LightBlue;       // Set background (fill) color
            bgStyle.ForegroundColor = Color.LightBlue;       // Foreground must match for solid fill

            // Apply the style to the whole sheet (covers all existing cells)
            sheet.Cells.ApplyStyle(bgStyle, new StyleFlag
            {
                All = true,               // Apply to all style attributes
                CellShading = true        // Ensure fill is applied
            });

            // -------------------------------------------------
            // 2) Set a background image for the worksheet
            // -------------------------------------------------
            // Path to the image file (ensure the file exists)
            string imagePath = "background.jpg";

            if (File.Exists(imagePath))
            {
                // Read the image into a byte array
                byte[] imageData = File.ReadAllBytes(imagePath);

                // Assign the image data to the worksheet's BackgroundImage property
                sheet.BackgroundImage = imageData;
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}");
            }

            // Save the workbook (lifecycle: save)
            string outputPath = "SheetWithBackground.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}