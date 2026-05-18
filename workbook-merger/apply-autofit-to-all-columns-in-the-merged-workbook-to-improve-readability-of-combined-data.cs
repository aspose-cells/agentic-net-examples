using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AutoFitColumnsAfterCombineDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                string sourcePath = "SourceWorkbook.xlsx";
                string secondPath = "SecondWorkbook.xlsx";
                string outputPath = "MergedAutoFitWorkbook.xlsx";

                // Ensure input files exist
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source workbook not found: {sourcePath}");
                if (!File.Exists(secondPath))
                    throw new FileNotFoundException($"Second workbook not found: {secondPath}");

                // Load workbooks
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Workbook secondWorkbook = new Workbook(secondPath);

                // Combine the second workbook into the source workbook
                sourceWorkbook.Combine(secondWorkbook);

                // AutoFit columns in each worksheet
                foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                {
                    sheet.AutoFitColumns();
                }

                // Save the merged workbook
                sourceWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Merged workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Preserve original Run method for compatibility
        public static void Run()
        {
            Main(null);
        }
    }
}