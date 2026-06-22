using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "source.xlsb";
            string outputPath = "output.xlsb";

            // Verify that the source file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the mashup data which contains Power Query formulas
            var mashup = workbook.DataMashup;

            // Check for Power Query formulas
            if (mashup?.PowerQueryFormulas != null && mashup.PowerQueryFormulas.Count > 0)
            {
                Console.WriteLine("Power Query Formulas found:");
                foreach (var formula in mashup.PowerQueryFormulas)
                {
                    Console.WriteLine($"Name       : {formula.Name}");
                    Console.WriteLine($"Definition : {formula.FormulaDefinition}");
                    Console.WriteLine($"Type       : {formula.Type}");
                    Console.WriteLine(new string('-', 30));
                }
            }
            else
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (no modifications made)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}