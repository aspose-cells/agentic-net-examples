using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    public class PowerQueryFormulaNamesDemo
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
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access DataMashup which holds Power Query formulas
            DataMashup mashupData = workbook.DataMashup;

            // Check for Power Query formulas
            if (mashupData?.PowerQueryFormulas != null && mashupData.PowerQueryFormulas.Count > 0)
            {
                Console.WriteLine("Power Query Formula Names:");
                foreach (PowerQueryFormula formula in mashupData.PowerQueryFormulas)
                {
                    Console.WriteLine(formula.Name);
                }
            }
            else
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
            }

            // Save the workbook (even if unchanged)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}