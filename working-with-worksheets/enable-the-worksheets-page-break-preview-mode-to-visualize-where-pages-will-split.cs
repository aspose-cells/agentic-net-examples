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

                // Enable Page Break Preview mode
                worksheet.IsPageBreakPreview = true;

                // Optional: set zoom to 100% for clearer view
                worksheet.Zoom = 100;

                // Output the current settings to the console
                Console.WriteLine("IsPageBreakPreview: " + worksheet.IsPageBreakPreview);
                Console.WriteLine("Zoom: " + worksheet.Zoom);

                // Define output file path
                string outputPath = "EnablePageBreakPreview_output.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
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