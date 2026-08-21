// Title: C# – Load Excel from MemoryStream, change a formula, and recalculate with CalculationOptions using Aspose.Cells
// Description: Demonstrates how to read an XLSX file into a byte array, open it with Aspose.Cells via a MemoryStream, update a cell formula, apply CalculationOptions (e.g., IgnoreError), recalculate all formulas with Workbook.CalculateFormula, and save the result back to a stream or file.
// Keywords: Aspose.Cells | C# load workbook from MemoryStream | modify cell formula programmatically | CalculationOptions | Workbook.CalculateFormula | ignore calculation errors | save workbook to stream | in‑memory Excel processing
// Common Searches: Aspose.Cells load workbook from byte array | change Excel cell formula in C# | recalculate formulas with options Aspose.Cells | ignore errors during Aspose.Cells calculation | save modified Excel to MemoryStream C#
// Developer Intent: Open an Excel workbook from a memory stream, edit a cell's formula, recalculate all formulas with custom options, and persist the changes without using the file system.
// Use Cases: Retrieve a workbook stored as a BLOB, adjust formulas, recalculate while suppressing errors, and write the updated file back to the database. | Process uploaded Excel files in a web API: modify formulas based on user input, recalculate with specific options, and return the updated file as a stream. | Generate a workbook entirely in memory, set dynamic formulas, perform a calculation with tailored settings, and stream the final file to a client application.
// AI Prompts: Generate C# code that opens an Excel file from a byte array with Aspose.Cells, changes cell C5 to '=SUM(A1:A10)', and recalculates all formulas while ignoring errors. | Show how to use Aspose.Cells CalculationOptions to disable iterative calculation and recalculate a workbook loaded from a MemoryStream. | Explain how to save a workbook after formula recalculation directly to a MemoryStream and obtain the resulting byte array in C#.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaUpdate
{
    // Demonstrates how to read an XLSX file into a byte array, open it with Aspose.Cells via a MemoryStream, update a cell formula, apply CalculationOptions (e.g., IgnoreError), recalculate all formulas with Workbook.CalculateFormula, and save the result back to a stream or file.
    class Program
    {
        static void Main()
        {
            // Assume the workbook binary data is available in a byte array.
            // In a real scenario this could come from a database, network, etc.
            byte[] workbookData = File.ReadAllBytes("input.xlsx");

            // Load the workbook from a memory stream.
            using (MemoryStream inputStream = new MemoryStream(workbookData))
            {
                Workbook workbook = new Workbook(inputStream);

                // Change an existing formula (or set a new one) in the first worksheet.
                // Example: modify cell B1 to multiply the value of A1 by 3.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["B1"].Formula = "=A1*3";

                // Prepare calculation options.
                CalculationOptions calcOptions = new CalculationOptions
                {
                    // Example option: ignore errors during calculation.
                    IgnoreError = true
                };

                // Recalculate all formulas in the workbook using the specified options.
                workbook.CalculateFormula(calcOptions);

                // (Optional) Save the updated workbook to a new memory stream or file.
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    // For demonstration, write the result to a file.
                    File.WriteAllBytes("output.xlsx", outputStream.ToArray());
                }

                // Display a result to verify the calculation.
                Console.WriteLine("Calculated value in B1: " + sheet.Cells["B1"].Value);
            }
        }
    }
}
