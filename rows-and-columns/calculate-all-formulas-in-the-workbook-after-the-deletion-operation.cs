// Title: C# – Delete a Row and Recalculate All Formulas with Aspose.Cells
// Description: Loads an Excel workbook, removes the third row (index 2) from the first worksheet while updating cross‑sheet references, forces a full formula recalculation across the entire workbook, and saves the result.
// Keywords: Aspose.Cells | C# DeleteRow | CalculateFormula | update references | Excel row removal | recalculate workbook formulas | cross‑sheet formula update | .NET Excel manipulation
// Common Searches: Aspose.Cells delete row and recalculate formulas | C# remove Excel row and refresh formulas | How to update formula references after deleting rows in Aspose.Cells | CalculateFormula after row deletion .NET | DeleteRow with updateReference true Aspose.Cells
// Developer Intent: Programmatically delete a specific row and ensure every formula in the workbook reflects the new layout.
// Use Cases: Delete a header row in a sales report and automatically refresh totals, percentages, and charts. | Remove a data entry row from a financial model and keep linked calculations on other sheets accurate. | Clean up imported data by deleting empty rows while guaranteeing that all dependent formulas recalculate correctly.
// AI Prompts: Generate C# code that deletes row 5 in an Excel workbook using Aspose.Cells, updates all references, and calls CalculateFormula with proper error handling. | Show how to delete multiple consecutive rows and then recalculate only the worksheets that contain affected formulas in Aspose.Cells for .NET. | Explain the effect of the DeleteRow method’s updateReference parameter on cross‑sheet formulas and how CalculateFormula resolves any remaining inconsistencies.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, removes the third row (index 2) from the first worksheet while updating cross‑sheet references, forces a full formula recalculation across the entire workbook, and saves the result.
    public class CalculateFormulasAfterDeletion
    {
        public static void Main(string[] args)
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load an existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Delete the third row (index 2) and update references in other worksheets
            cells.DeleteRow(2, true);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
