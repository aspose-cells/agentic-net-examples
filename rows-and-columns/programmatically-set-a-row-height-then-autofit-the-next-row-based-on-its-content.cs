using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetRowHeightAndAutoFitNextRow
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Set a custom height (in points) for the first row (index 0)
                worksheet.Cells.SetRowHeight(0, 20); // 20 points

                // Populate the second row (index 1) with long text that will require auto‑fit
                worksheet.Cells["A2"].PutValue("This is a very long piece of text that should cause the row height to increase when auto‑fitted.");
                worksheet.Cells["B2"].PutValue("Additional long text in the same row to demonstrate auto‑fit behavior.");

                // Auto‑fit the second row based on its content
                worksheet.AutoFitRow(1);

                // Define output file path
                string outputPath = "SetRowHeightAndAutoFitNextRow.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRowHeightAndAutoFitNextRow.Run();
        }
    }
}