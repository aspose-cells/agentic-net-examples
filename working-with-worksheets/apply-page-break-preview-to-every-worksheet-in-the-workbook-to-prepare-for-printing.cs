using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ApplyPageBreakPreviewToAllSheets
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "Output_WithPageBreakPreview.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one sheet
            }

            // Enable Page Break Preview for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.IsPageBreakPreview = true;
                // Optional: sheet.ViewType = ViewType.PageBreakPreview;
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}