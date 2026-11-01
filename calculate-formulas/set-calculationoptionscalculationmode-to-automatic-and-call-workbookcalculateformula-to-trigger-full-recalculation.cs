// Title: Set Workbook.CalculationOptions.CalculationMode to Automatic and force full formula recalculation with Aspose.Cells for .NET
// Common Searches: how to enable automatic calculation mode in Aspose.Cells C# workbook | force full formula recalculation in an Excel file using Aspose.Cells CalculateFormula | Aspose.Cells set CalculationOptions.CalculationMode to Automatic before saving workbook | recalculate all formulas after loading an existing workbook with Aspose.Cells .NET | C# code to trigger workbook.CalculateFormula after changing calculation settings
// Tags: Aspose.Cells automatic calculation setting | Workbook.CalculateFormula example | full workbook formula recalculation .NET | C# set CalculationOptions mode | Excel file formula refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# example demonstrates how to ensure that an Excel workbook uses the Automatic calculation mode with Aspose.Cells, invoke Workbook.CalculateFormula to force a complete recalculation of every formula, and then save the updated file. The code loads an existing workbook (or creates a new one), optionally sets workbook.CalculationOptions.CalculationMode = CalculationMode.Automatic, calls CalculateFormula, and writes the result to output.xlsx while handling file‑not‑found and generic exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load existing workbook if the file exists; otherwise create a new workbook
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Recalculate all formulas in the workbook (default calculation mode is Automatic)
                workbook.CalculateFormula();

                // Save the workbook to the specified output file
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
