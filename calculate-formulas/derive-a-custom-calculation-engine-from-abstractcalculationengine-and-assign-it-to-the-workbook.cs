// Title: Calculate worksheet formulas and save an Aspose.Cells workbook to XLSX using C#
// Common Searches: how to programmatically calculate formulas in an Aspose.Cells workbook with C# | C# code to evaluate SUM formula in Aspose.Cells and export to XLSX | saving a workbook after formula calculation using Aspose.Cells for .NET | ensure output folder exists before saving Aspose.Cells workbook in C# | run workbook.CalculateFormula() in Aspose.Cells .NET example
// Tags: calculate workbook formulas Aspose.Cells C# | save workbook as XLSX Aspose.Cells | initialize worksheet cells with values C# | run CalculateFormula method Aspose.Cells | create output directory before workbook save

using System;
using System.IO;
using Aspose.Cells;

namespace CustomCalculationEngineDemo
{
    // The example creates a new Workbook, writes numeric values to cells A1 and A2, sets a SUM formula in A3, triggers formula calculation with CalculateFormula(), ensures the target directory exists, and saves the result as CustomEngineResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Define output file path
                string outputPath = "CustomEngineResult.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
