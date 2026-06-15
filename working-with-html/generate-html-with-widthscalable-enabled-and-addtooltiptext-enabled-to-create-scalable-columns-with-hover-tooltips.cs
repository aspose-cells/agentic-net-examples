using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HtmlSaveOptionsScalableTooltipDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data that may be truncated in a narrow column
                worksheet.Cells["A1"].PutValue("This is a very long text that will not fit in the column width and should show a tooltip.");
                worksheet.Cells.SetColumnWidth(0, 10); // Narrow column to force truncation

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    WidthScalable = true,      // Enable scalable column widths
                    AddTooltipText = true      // Enable tooltip for truncated content
                };

                // Define output path
                string outputPath = "ScalableWithTooltip.html";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML with the specified options
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved to '{outputPath}' with WidthScalable and AddTooltipText enabled.");
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
            HtmlSaveOptionsScalableTooltipDemo.Run();
        }
    }
}