using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadFirstThreeSheets
    {
        public static void Run()
        {
            try
            {
                // Path to the source workbook
                string sourcePath = "input.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook (all sheets)
                Workbook workbook = new Workbook(sourcePath);

                // Keep only the first three worksheets
                while (workbook.Worksheets.Count > 3)
                {
                    // Remove sheets starting from index 3 (the fourth sheet)
                    workbook.Worksheets.RemoveAt(3);
                }

                // Display loaded worksheets
                Console.WriteLine("Number of worksheets loaded: " + workbook.Worksheets.Count);
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Console.WriteLine($"Sheet {i}: {workbook.Worksheets[i].Name}");
                }

                // Save the partially loaded workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadFirstThreeSheets.Run();
        }
    }
}