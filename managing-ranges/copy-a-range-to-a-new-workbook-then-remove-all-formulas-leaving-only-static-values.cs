// Title: Copy a range to a new workbook and strip formulas with Aspose.Cells for .NET
// Description: Load a source workbook, copy a defined range (e.g., A1:D10) to a fresh workbook using PasteOptions.Values, then call RemoveFormulas to keep only static values before saving.
// Keywords: Aspose.Cells | C# | copy range values only | remove formulas | PasteOptions.Values | Excel static values | read‑only report workbook | archive snapshot Excel
// Common Searches: Aspose.Cells copy range without formulas | C# copy cells as values to new workbook | remove all formulas after copying Excel file Aspose | PasteOptions.Values example Aspose.Cells | how to export static values from Excel using Aspose
// Developer Intent: Transfer a specific cell block from one workbook to another while ensuring the destination contains only literal values, no formulas.
// Use Cases: Generate a read‑only report by copying calculated results as plain numbers. | Create an archival snapshot of a worksheet that hides underlying formulas. | Distribute a pre‑filled template to external partners without exposing calculation logic.
// AI Prompts: Show C# code that copies a range from one workbook to another with Aspose.Cells and retains only the values. | Demonstrate how to purge any remaining formulas after copying cells using Aspose.Cells. | Explain the effect of PasteOptions.Values when copying ranges between Excel workbooks in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCopyRangeAndRemoveFormulas
{
    // Load a source workbook, copy a defined range (e.g., A1:D10) to a fresh workbook using PasteOptions.Values, then call RemoveFormulas to keep only static values before saving.
    class Program
    {
        static void Main()
        {
            // Paths for source and destination workbooks
            string sourcePath = "source.xlsx";
            string destinationPath = "destination.xlsx";

            try
            {
                // Ensure source workbook exists; create a simple one if missing
                if (!File.Exists(sourcePath))
                {
                    var tempWb = new Workbook();
                    var tempWs = tempWb.Worksheets[0];
                    // Populate some sample data
                    tempWs.Cells["A1"].PutValue("Item");
                    tempWs.Cells["B1"].PutValue(123);
                    // Set a formula in C1
                    tempWs.Cells["C1"].Formula = "=B1*2";
                    tempWb.Save(sourcePath);
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new (empty) workbook for the destination
                Workbook destinationWorkbook = new Workbook();

                // Get the first worksheet from each workbook
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

                // Define the range to copy (adjust the address as needed)
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:D10");

                // Create a matching range in the destination worksheet
                AsposeRange destinationRange = destinationSheet.Cells.CreateRange("A1:D10");

                // Set paste options to copy only the values (no formulas)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.Values
                };

                // Copy the source range to the destination range using the paste options
                destinationRange.Copy(sourceRange, pasteOptions);

                // As an extra safeguard, remove any formulas that might still exist
                destinationSheet.Cells.RemoveFormulas();

                // Ensure the destination directory exists
                string destDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the new workbook
                destinationWorkbook.Save(destinationPath);
                Console.WriteLine($"Workbook copied successfully to '{destinationPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
