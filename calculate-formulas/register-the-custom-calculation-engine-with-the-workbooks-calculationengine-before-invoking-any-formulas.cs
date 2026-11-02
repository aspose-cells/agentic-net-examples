// Title: How to register a custom calculation engine with an Aspose.Cells workbook in C# before evaluating formulas
// Common Searches: c# set custom calculation engine for Aspose.Cells workbook before CalculateFormula | how to replace default formula evaluator with a custom engine in Aspose.Cells | Aspose.Cells custom calculation engine example C# | register user-defined calculation engine in an Aspose.Cells workbook | override Excel formula calculation in Aspose.Cells using a custom engine
// Tags: custom calculation engine registration Aspose.Cells | override default formula evaluator C# | Aspose.Cells workbook calculation engine customization | replace Aspose.Cells formula evaluator with custom engine | Excel formula calculation custom engine Aspose.Cells | assign custom engine before CalculateFormula

using System;
using System.IO;
using Aspose.Cells;

namespace CustomCalculationEngineDemo
{
    // The sample creates a new Workbook, writes values to cells A1 and A2, sets a SUM formula in A3, and saves the file as Output.xlsx. To use a custom calculation engine, assign an implementation of ICalculationEngine to workbook.CalculationEngine before calling workbook.CalculateFormula, which replaces the default evaluator for all subsequent formula calculations.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (create rule)
                Workbook workbook = new Workbook();

                // Example: add data and a formula
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

                // Invoke calculation (uses default engine)
                workbook.CalculateFormula();

                // Define output path and ensure directory exists
                string outputPath = "Output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (save rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
