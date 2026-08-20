// Title: Expand a Named Range with Name.RefersTo and Recalculate Formulas using Aspose.Cells for .NET
// Description: Loads an Excel workbook, ensures a named range (e.g., "MyRange") exists, updates its RefersTo property to a larger area (A1:C10), forces a full formula recalculation with CalculateFormula, and saves the result.
// Keywords: Aspose.Cells | C# | named range | Name.RefersTo | expand range | recalculate formulas | Workbook.CalculateFormula | create missing name | Excel automation
// Common Searches: Aspose.Cells change RefersTo of a named range | expand named range and recalc formulas .NET | create named range if not exists Aspose.Cells | update named range address C# | force formula recalculation after range change
// Developer Intent: Modify an existing named range to cover a larger cell block and refresh all dependent formulas in the workbook.
// Use Cases: Enlarge a data range used by a chart after adding new rows. | Extend a validation range before importing additional records. | Adjust a range referenced by SUM/AVERAGE formulas when extra columns are added.
// AI Prompts: Generate C# code that checks for a named range, creates it if missing, sets RefersTo to a new address, and calls CalculateFormula with Aspose.Cells. | Explain the correct RefersTo string format for updating a named range in Aspose.Cells. | Provide a step‑by‑step tutorial to replace a named range with a larger area and ensure all dependent formulas are recalculated.

using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceNamedRangeExample
{
    // Loads an Excel workbook, ensures a named range (e.g., "MyRange") exists, updates its RefersTo property to a larger area (A1:C10), forces a full formula recalculation with CalculateFormula, and saves the result.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (assumed to contain the named range)
                Worksheet sheet = workbook.Worksheets[0];

                // Name of the existing named range to be replaced
                string existingName = "MyRange";

                // Retrieve the Name object; create it if it does not exist
                Name namedRange = workbook.Worksheets.Names[existingName];
                if (namedRange == null)
                {
                    // Add returns the index of the newly created name
                    int index = workbook.Worksheets.Names.Add(existingName);
                    namedRange = workbook.Worksheets.Names[index];
                }

                // Define the new, larger area for the named range (e.g., A1:C10 on the same sheet)
                // The RefersTo string must start with an equal sign and include the sheet name.
                string newRefersTo = $"={sheet.Name}!$A$1:$C$10";

                // Update the RefersTo property to point to the new area
                namedRange.RefersTo = newRefersTo;

                // Recalculate all formulas in the workbook so they reflect the updated range
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
