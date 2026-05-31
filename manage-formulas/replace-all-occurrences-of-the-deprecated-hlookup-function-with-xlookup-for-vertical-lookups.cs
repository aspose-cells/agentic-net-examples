using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ReplaceHlookupWithXlookup
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);

                // Replace deprecated HLOOKUP with XLOOKUP in formulas
                workbook.Replace("HLOOKUP", "XLOOKUP");

                // Recalculate formulas
                workbook.CalculateFormula();

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceHlookupWithXlookup.Run();
        }
    }
}