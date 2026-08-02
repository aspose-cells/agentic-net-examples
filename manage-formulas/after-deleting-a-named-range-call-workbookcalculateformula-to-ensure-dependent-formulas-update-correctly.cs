// Title: Aspose.Cells for .NET – Delete a Named Range and Recalculate Formulas (C#)
// Description: Demonstrates how to create a workbook, define a named range, use it in a SUM formula, delete the range, and invoke workbook.CalculateFormula to refresh dependent cells before saving the file.
// Keywords: Aspose.Cells | C# | .NET | named range deletion | remove named range programmatically | recalculate formulas | CalculateFormula method | Excel automation | update dependent formulas | SUM formula with named range | workbook.Save
// Common Searches: Aspose.Cells delete named range C# | How to refresh formulas after removing a named range in Aspose.Cells | CalculateFormula after name removal .NET | Update Excel formulas when named range is deleted using Aspose | C# example for removing named ranges and recalculating workbook
// Developer Intent: Remove a defined name from a workbook and ensure all formulas that referenced it are recalculated automatically.
// Use Cases: Cleaning up temporary named ranges before exporting a report while preserving accurate totals. | Automating spreadsheet generation where dynamic named ranges are created and later discarded, requiring formula updates. | Implementing user‑driven edits that delete named ranges and instantly reflect the changes in dependent calculations.
// AI Prompts: Show me C# code to delete a named range in Aspose.Cells and recalculate affected formulas. | Explain why calling workbook.CalculateFormula is required after removing a named range. | Provide an example that removes multiple named ranges and updates all related formulas in a .NET workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a named range, use it in a SUM formula, delete the range, and invoke workbook.CalculateFormula to refresh dependent cells before saving the file.
    public class DeleteNamedRangeAndRecalculate
    {
        // Entry point for the application
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data in column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Create a named range that refers to A1:A3
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

            // Use the named range in a formula
            sheet.Cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate formulas before deletion (optional, just to have initial values)
            workbook.CalculateFormula();

            // Delete the named range
            workbook.Worksheets.Names.Remove("MyRange");

            // Recalculate formulas so that dependent cells update correctly
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("DeletedNamedRange.xlsx");
        }
    }
}
