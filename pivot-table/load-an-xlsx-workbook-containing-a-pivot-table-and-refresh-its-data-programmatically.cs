using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class RefreshPivotTableDemo
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Refresh all PivotTables in every worksheet
                workbook.Worksheets.RefreshPivotTables();

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshPivotTableDemo.Run();
        }
    }
}