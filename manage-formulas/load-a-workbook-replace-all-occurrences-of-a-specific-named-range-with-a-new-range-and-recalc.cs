// Title: C# – Replace a Named Range and Recalculate Formulas with Aspose.Cells for .NET
// Description: Loads an Excel workbook, finds a named range by its identifier, changes its RefersTo address to a new range, triggers a full formula recalculation, and saves the updated file using Aspose.Cells.
// Keywords: Aspose.Cells | C# | named range replace | RefersTo property | formula recalculation | Excel automation | update named range address | Workbook.CalculateFormula | Aspose.Cells API | Excel workbook modification
// Common Searches: Aspose.Cells change named range address C# | update RefersTo property Aspose.Cells | recalculate formulas after named range edit | find and replace named range in Excel using Aspose.Cells | C# code to modify named range and recalc workbook
// Developer Intent: Replace an existing named range with a new address and refresh all dependent formulas.
// Use Cases: Adjust data source ranges in financial models without breaking formulas. | Redirect chart series to a new cell block after restructuring a worksheet. | Migrate legacy named ranges to a new layout across a batch of workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to locate a named range, set a new RefersTo address, recalculate all formulas, and save the workbook. | Explain the impact of changing a named range's RefersTo value on existing formulas in Aspose.Cells and how to ensure accurate recalculation.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeReplaceDemo
{
    // Loads an Excel workbook, finds a named range by its identifier, changes its RefersTo address to a new range, triggers a full formula recalculation, and saves the updated file using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // The name of the existing named range to be replaced
            string oldRangeName = "OldRange";

            // The address of the new range (example: Sheet1!$B$1:$B$5)
            // Adjust the sheet name and address as needed
            string newRangeAddress = "=Sheet1!$B$1:$B$5";

            // Locate the Name object that represents the old named range
            Name targetName = null;
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (name.Text.Equals(oldRangeName, StringComparison.OrdinalIgnoreCase))
                {
                    targetName = name;
                    break;
                }
            }

            if (targetName == null)
            {
                Console.WriteLine($"Named range '{oldRangeName}' not found.");
                return;
            }

            // Update the RefersTo property to point to the new range
            targetName.RefersTo = newRangeAddress;

            // Recalculate all formulas so that any references to the renamed range are updated
            workbook.CalculateFormula();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");

            Console.WriteLine("Named range replacement completed successfully.");
        }
    }
}
