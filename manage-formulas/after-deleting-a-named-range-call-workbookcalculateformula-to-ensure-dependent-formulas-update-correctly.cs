// Title: Aspose.Cells .NET: Delete a Named Range and Recalculate Formulas
// Description: C# example that creates a workbook, defines a named range, uses it in a SUM formula, removes the named range, calls Workbook.CalculateFormula to refresh dependent cells, and saves the file. Demonstrates proper cleanup of named ranges and formula updates with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | delete named range | Workbook.CalculateFormula | update dependent formulas | named range removal | Excel automation | Aspose.Cells API
// Common Searches: Aspose.Cells delete named range C# | How to refresh formulas after removing a named range in Aspose.Cells | Workbook.CalculateFormula after name removal | C# code to remove named range and recalculate | Aspose.Cells named range cleanup
// Developer Intent: Remove a specific named range from a workbook and trigger a full formula recalculation so that any cells referencing the range reflect the change.
// Use Cases: Eliminate temporary named ranges before publishing a report | Automate cleanup of legacy named ranges in generated spreadsheets | Ensure summary totals stay accurate after programmatically deleting a range | Prepare workbooks for third‑party systems that do not support custom names
// AI Prompts: Write C# code using Aspose.Cells to delete a named range called 'MyRange' and then call Workbook.CalculateFormula to update all dependent cells. | Show how to handle the case where the named range might not exist before removal, with proper exception handling. | Explain the impact of Workbook.CalculateFormula on performance when called after multiple named‑range deletions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, defines a named range, uses it in a SUM formula, removes the named range, calls Workbook.CalculateFormula to refresh dependent cells, and saves the file. Demonstrates proper cleanup of named ranges and formula updates with Aspose.Cells for .NET.
    public class DeleteNamedRangeAndRecalculate
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data in cells A1:A3
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Create a named range "MyRange" that refers to A1:A3
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

                // Use the named range in a formula (sum of the range)
                sheet.Cells["D1"].Formula = "=SUM(MyRange)";

                // Calculate formulas before removing the named range (optional)
                workbook.CalculateFormula();

                // Remove the named range from the workbook
                workbook.Worksheets.Names.Remove("MyRange");

                // Recalculate formulas so that dependent cells update correctly
                workbook.CalculateFormula();

                // Define output file path
                string outputPath = "DeletedNamedRange.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
