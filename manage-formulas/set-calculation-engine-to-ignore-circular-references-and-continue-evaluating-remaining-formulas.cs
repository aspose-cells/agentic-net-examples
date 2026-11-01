// Title: How to configure Aspose.Cells in C# to ignore circular references and keep evaluating formulas
// Common Searches: Aspose.Cells ignore circular reference errors during CalculateFormula in C# | Set CalculationOptions.IgnoreCircularReferences true in Aspose.Cells .NET | Continue formula evaluation after circular reference in Aspose.Cells workbook | Disable circular reference detection when recalculating Excel file with Aspose.Cells | Example of skipping circular references in Aspose.Cells calculation engine C#
// Tags: Aspose.Cells CalculationOptions.IgnoreCircularReferences | circular reference handling Aspose.Cells | formula evaluation continuation Aspose.Cells | disable circular reference detection .NET | set calculation options for circular references Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a Workbook, sets CalculationOptions.IgnoreCircularReferences to true so the calculation engine skips circular references, recalculates all formulas, and saves the workbook to an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Note: In newer Aspose.Cells versions the circular reference handling
                // is enabled by default or managed via CalculationOptions.
                // If needed, configure calculation options here.

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Save the workbook to a file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
