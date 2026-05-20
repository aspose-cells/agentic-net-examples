using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsDemo
{
    public class SparklineToStreamExample
    {
        // Renders a sparkline to a memory stream (PNG) for later embedding into a PDF.
        public static MemoryStream RenderSparklineToStream()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that the sparkline will represent.
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["B1"].PutValue(20);
                sheet.Cells["C1"].PutValue(15);
                sheet.Cells["D1"].PutValue(30);

                // Define the cell where the sparkline will be placed.
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a sparkline group (Line type) and retrieve the first sparkline.
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];
                Sparkline sparkline = group.Sparklines[0];

                // Configure image options for the output image.
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    Transparent = false
                };

                // Render the sparkline into a memory stream.
                MemoryStream stream = new MemoryStream();
                sparkline.ToImage(stream, options);
                stream.Position = 0; // Reset stream position for subsequent reading.

                return stream;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error rendering sparkline: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                using (MemoryStream sparklineStream = SparklineToStreamExample.RenderSparklineToStream())
                {
                    // Save the stream to a file to verify the output.
                    string outputPath = "sparkline.png";
                    using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        sparklineStream.CopyTo(file);
                    }
                    Console.WriteLine($"Sparkline image saved to {Path.GetFullPath(outputPath)}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}