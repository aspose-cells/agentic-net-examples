using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class WorkbookLoadSaveDemo
    {
        public static void Main(string[] args)
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook inside a using block for deterministic disposal
            using (Workbook workbook = new Workbook(inputPath))
            {
                // Modify cell A1
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Modified");

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}