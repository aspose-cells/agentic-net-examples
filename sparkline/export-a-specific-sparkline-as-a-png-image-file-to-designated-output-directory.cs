using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace SparklineExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].PutValue(15);
            sheet.Cells["D1"].PutValue(30);

            // Define the cell where the sparkline will be placed
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group (Line type) and retrieve the first sparkline
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, sparklineLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];
            Sparkline sparkline = group.Sparklines[0];

            // Configure image options for PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                Transparent = false
            };

            // Ensure the output directory exists
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // Build the full file path for the exported sparkline image
            string imagePath = Path.Combine(outputDir, "sparkline.png");

            // Export the sparkline to a PNG file
            sparkline.ToImage(imagePath, imgOptions);

            Console.WriteLine($"Sparkline image saved to: {imagePath}");
        }
    }
}