using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RefreshConnectionsAndFreezeRows
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all data connections, pivot tables, and charts
            workbook.RefreshAll();

            // Freeze the top row in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // Freeze first row (row index 1) with no columns frozen
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}