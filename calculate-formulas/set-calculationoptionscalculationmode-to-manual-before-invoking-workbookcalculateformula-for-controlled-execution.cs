// Title: How to set CalculationOptions.CalculationMode to Manual before calling Workbook.CalculateFormula in Aspose.Cells for .NET
// Common Searches: Aspose.Cells set CalculationOptions.CalculationMode to Manual before Workbook.CalculateFormula | C# disable automatic formula recalculation in Aspose.Cells | manual calculation mode Aspose.Cells .NET example | control when formulas are evaluated with Aspose.Cells Workbook.CalculateFormula | prevent auto-recalc when loading workbook using Aspose.Cells C#
// Tags: Aspose.Cells calculation mode manual | Workbook.CalculateFormula manual execution | set CalculationOptions.CalculationMode C# | prevent auto formula recalculation Aspose.Cells | controlled formula evaluation .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook or creates a new one, sets CalculationOptions.CalculationMode to Manual to stop automatic recalculation, explicitly runs Workbook.CalculateFormula, and then saves the workbook. This ensures formulas are evaluated only when you intend.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file names
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            Workbook workbook;

            // Load an existing workbook if the file exists; otherwise create a new one
            if (File.Exists(inputFile))
            {
                workbook = new Workbook(inputFile);
            }
            else
            {
                workbook = new Workbook();
            }

            // Perform formula calculation explicitly
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Formula calculation error: {calcEx.Message}");
            }

            // Save the workbook after calculation
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
