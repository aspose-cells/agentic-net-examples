using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AutoFitRowsRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate rows 15‑20 (zero‑based indices 14‑19) with sample data
                for (int row = 14; row <= 19; row++)
                {
                    // Add some long text to column A to demonstrate row height adjustment
                    worksheet.Cells[row, 0].PutValue(
                        $"This is a long text in row {row + 1} that should cause the row to expand when auto‑fitted.");

                    // Enable text wrapping so the height can increase
                    Style style = worksheet.Cells[row, 0].GetStyle();
                    style.IsTextWrapped = true;
                    worksheet.Cells[row, 0].SetStyle(style);
                }

                // Auto‑fit rows 15‑20 (indices 14‑19)
                worksheet.AutoFitRows(14, 19);

                // Determine output path (Desktop folder)
                string outputPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "AutoFitRowsRangeDemo.xlsx");

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to: {outputPath}");
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
            AutoFitRowsRangeDemo.Run();
        }
    }
}