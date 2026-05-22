using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertCommaSeparatedColumnB
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate column B (index 1) with comma‑separated values
                cells["B1"].PutValue("John,Doe,30");
                cells["B2"].PutValue("Jane,Smith,28");
                cells["B3"].PutValue("Bob,Johnson,45");

                // Configure TextToColumns options to use comma as the delimiter
                TxtLoadOptions options = new TxtLoadOptions
                {
                    Separator = ',' // comma delimiter
                };

                // Apply TextToColumns starting from row 0, column 1 (B), processing 3 rows
                // This will split each cell's content into separate columns (C, D, etc.)
                cells.TextToColumns(0, 1, 3, options);

                // Determine output path and save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedColumnB.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertCommaSeparatedColumnB.Run();
        }
    }
}