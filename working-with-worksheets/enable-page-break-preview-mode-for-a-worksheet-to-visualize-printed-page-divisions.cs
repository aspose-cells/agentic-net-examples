using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EnablePageBreakPreview
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable page break preview mode
                worksheet.IsPageBreakPreview = true;

                // Set zoom to 100% for clearer view (optional)
                worksheet.Zoom = 100;

                // Define output file path
                string outputPath = "PageBreakPreviewDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            EnablePageBreakPreview.Run();
        }
    }
}