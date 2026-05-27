using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SwitchToNormalViewDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook or load an existing one if the file is present
                Workbook workbook;
                const string inputPath = "input.xlsx";

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Switch the worksheet back to Normal view
                worksheet.IsPageBreakPreview = false;
                // Alternative: worksheet.ViewType = ViewType.NormalView;

                // Verify the current view mode
                Console.WriteLine("IsPageBreakPreview: " + worksheet.IsPageBreakPreview);
                Console.WriteLine("ViewType: " + worksheet.ViewType);

                // Save the workbook
                const string outputPath = "SwitchToNormalViewDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}