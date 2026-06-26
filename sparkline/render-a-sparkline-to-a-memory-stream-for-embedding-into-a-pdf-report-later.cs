using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace SparklineDemo
{
    public class SparklineToMemoryStreamDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["B1"].PutValue(20);
                sheet.Cells["C1"].PutValue(15);
                sheet.Cells["D1"].PutValue(30);

                // Define the cell area where the sparkline will be placed (E1)
                CellArea sparklineLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a sparkline group with the data range A1:D1
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, sparklineLocation);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Retrieve the first sparkline from the group
                Sparkline sparkline = group.Sparklines[0];

                // Configure image options for the sparkline rendering
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    Quality = 90,
                    Transparent = false
                };

                // Render the sparkline to a memory stream
                using (MemoryStream stream = new MemoryStream())
                {
                    sparkline.ToImage(stream, options);
                    stream.Position = 0;
                    Console.WriteLine($"Sparkline image rendered to memory stream. Length = {stream.Length} bytes.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SparklineToMemoryStreamDemo.Run();
        }
    }
}