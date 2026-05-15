using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSparklineExport
{
    public class ExportSparklineAsPng
    {
        public static void Run()
        {
            // Define output directory and ensure it exists
            string outputDir = Path.Combine(Environment.CurrentDirectory, "SparklineImages");
            Directory.CreateDirectory(outputDir);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].PutValue(15);
            sheet.Cells["D1"].PutValue(30);

            // Define the location where the sparkline will be placed (E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group of type Line with the data range and location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Retrieve the first sparkline from the group
            Sparkline sparkline = group.Sparklines[0];

            // Configure image options for PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Build the full file path for the exported image
            string imagePath = Path.Combine(outputDir, "sparkline_output.png");

            // Export the sparkline to PNG using the ToImage method (string, ImageOrPrintOptions)
            sparkline.ToImage(imagePath, imgOptions);

            Console.WriteLine($"Sparkline exported successfully to: {imagePath}");

            // Optionally, save the workbook for reference
            string workbookPath = Path.Combine(outputDir, "WorkbookWithSparkline.xlsx");
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to: {workbookPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportSparklineAsPng.Run();
        }
    }
}