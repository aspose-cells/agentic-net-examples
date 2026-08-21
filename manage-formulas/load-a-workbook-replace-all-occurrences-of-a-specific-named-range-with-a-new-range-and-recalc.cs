// Title: Replace a Named Range and Recalculate Formulas with Aspose.Cells for .NET
// Description: Loads a workbook, changes the RefersTo address of a specified named range, triggers a full formula recalculation, and saves the updated file.
// Keywords: Aspose.Cells replace named range | update RefersTo address .NET | recalculate formulas Aspose.Cells | modify named range programmatically | C# Aspose.Cells workbook edit
// Common Searches: how to change a named range address using Aspose.Cells | recalculate all formulas after updating a named range .NET | Aspose.Cells replace OldRange with Sheet1!B1:B5 | C# code to edit named ranges in Excel files
// Developer Intent: Change an existing named range to a new cell range and refresh all dependent formulas.
// Use Cases: Redirect a named range after data migration so existing formulas point to the new column. | Adjust a range reference when rows are inserted, keeping financial models accurate. | Swap a range before generating a report to ensure calculations use the latest dataset.
// AI Prompts: Write C# that verifies a named range exists, creates it if missing, and sets its RefersTo address with Aspose.Cells. | Show how to update several named ranges in one workbook and then call CalculateFormula to refresh dependent cells. | Explain error handling for invalid RefersTo strings when modifying a Name object in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsReplaceNamedRange
{
    // Loads a workbook, changes the RefersTo address of a specified named range, triggers a full formula recalculation, and saves the updated file.
    public class Program
    {
        public static void Main()
        {
            // Path to the existing workbook
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Name of the existing named range to be replaced
            string oldNamedRange = "OldRange";

            // New range address (including sheet name) that will replace the old named range
            // Example: replace with cells B1:B5 on Sheet1
            string newRangeAddress = "=Sheet1!$B$1:$B$5";

            // Retrieve the Name object for the old named range
            Name name = workbook.Worksheets.Names[oldNamedRange];

            if (name != null)
            {
                // Update the RefersTo property to point to the new range
                name.RefersTo = newRangeAddress;
            }
            else
            {
                Console.WriteLine($"Named range '{oldNamedRange}' not found.");
                return;
            }

            // Recalculate all formulas to reflect the changed named range
            workbook.CalculateFormula();

            // Save the modified workbook
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}
