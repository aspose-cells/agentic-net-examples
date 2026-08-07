// Title: Load Workbook from MemoryStream, Change Formula, Recalculate with CalculationOptions – Aspose.Cells C#
// Description: Demonstrates how to create an Excel workbook, save it to a MemoryStream, reload it, replace the formula in cell B1, set custom CalculationOptions (Recursive = true, IgnoreError = false), and invoke Workbook.CalculateFormula to obtain the updated result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells MemoryStream | C# load workbook from stream | modify cell formula Aspose.Cells | Workbook.CalculateFormula options | CalculationOptions Recursive | ignore errors Aspose.Cells | in‑memory Excel processing | .NET Excel formula recalculation
// Common Searches: Aspose.Cells load workbook from MemoryStream C# | change Excel cell formula programmatically Aspose.Cells | calculate formulas with custom options Aspose.Cells | recursive calculation option Aspose.Cells | how to recalculate all formulas after editing a workbook
// Developer Intent: Load an Excel file from a MemoryStream, update a cell's formula, and recalculate all formulas using specific CalculationOptions in Aspose.Cells for .NET.
// Use Cases: Process uploaded Excel files in a web API without writing to disk, adjust formulas, and return calculated values. | Perform bulk formula modifications in in‑memory workbooks for financial or scientific models before saving or exporting. | Enable recursive calculations across multiple worksheets after programmatically changing formulas in a streamed workbook.
// AI Prompts: Generate C# code that loads an Excel workbook from a MemoryStream, changes B1's formula to =A1*A2, sets CalculationOptions (Recursive = true, IgnoreError = false), and runs Workbook.CalculateFormula. | Show how to configure Aspose.Cells CalculationOptions to enable recursive calculation and prevent error suppression when recalculating formulas after a formula change. | Provide an example of reading an XLSX file into a MemoryStream, updating a formula, recalculating with custom options, and retrieving the new cell value using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to create an Excel workbook, save it to a MemoryStream, reload it, replace the formula in cell B1, set custom CalculationOptions (Recursive = true, IgnoreError = false), and invoke Workbook.CalculateFormula to obtain the updated result using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook and save it into a memory stream
            // ------------------------------------------------------------
            Workbook originalWorkbook = new Workbook();
            Worksheet sheet = originalWorkbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put some initial values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Set an initial formula that will be changed later
            cells["B1"].Formula = "=A1+A2";

            // Save the workbook to a memory stream (in XLSX format)
            using (MemoryStream memStream = new MemoryStream())
            {
                originalWorkbook.Save(memStream, SaveFormat.Xlsx);

                // Reset the stream position to the beginning for reading
                memStream.Position = 0;

                // ------------------------------------------------------------
                // 2. Load the workbook from the memory stream
                // ------------------------------------------------------------
                Workbook loadedWorkbook = new Workbook(memStream);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                Cells loadedCells = loadedSheet.Cells;

                // ------------------------------------------------------------
                // 3. Change the formula in cell B1
                // ------------------------------------------------------------
                // New formula: multiply A1 and A2 instead of adding them
                loadedCells["B1"].Formula = "=A1*A2";

                // ------------------------------------------------------------
                // 4. Prepare calculation options
                // ------------------------------------------------------------
                CalculationOptions calcOptions = new CalculationOptions
                {
                    // Example: enable recursive calculation across worksheets (if any)
                    Recursive = true,
                    // Example: do not ignore errors during calculation
                    IgnoreError = false
                };

                // ------------------------------------------------------------
                // 5. Recalculate all formulas using the specified options
                // ------------------------------------------------------------
                loadedWorkbook.CalculateFormula(calcOptions);

                // ------------------------------------------------------------
                // 6. Output the result of the changed formula
                // ------------------------------------------------------------
                Console.WriteLine("Result of B1 after recalculation: " + loadedCells["B1"].Value);
            }
        }
    }
}
