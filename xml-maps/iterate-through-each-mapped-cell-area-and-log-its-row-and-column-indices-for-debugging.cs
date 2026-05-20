using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class IterateMappedCellAreas
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: create some merged cell areas to have mapped CellArea objects
            cells.Merge(0, 0, 1, 1); // Merge A1:B2
            cells.Merge(2, 2, 3, 3); // Merge C3:D4

            // Retrieve all merged cell areas (mapped cell areas)
            CellArea[] mappedAreas = cells.GetMergedAreas();

            // Iterate through each CellArea and log its row/column indices
            foreach (CellArea area in mappedAreas)
            {
                Console.WriteLine(
                    $"Mapped Area - StartRow: {area.StartRow}, StartColumn: {area.StartColumn}, " +
                    $"EndRow: {area.EndRow}, EndColumn: {area.EndColumn}");
            }

            // Define output file path
            string outputPath = "IterateMappedCellAreas.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}