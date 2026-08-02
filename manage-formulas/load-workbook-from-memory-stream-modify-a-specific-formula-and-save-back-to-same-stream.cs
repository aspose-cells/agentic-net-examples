// Title: C# – Load, Edit a Formula, and Save an Excel Workbook in the Same MemoryStream with Aspose.Cells
// Description: Demonstrates how to load an XLSX workbook from a MemoryStream, change a cell formula, recalculate, clear the stream, and write the updated workbook back to the original MemoryStream using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | MemoryStream | load workbook from stream | modify Excel formula | save workbook to same stream | recalculate formulas | in‑memory Excel processing | byte array Excel update
// Common Searches: Aspose.Cells load workbook from MemoryStream C# | change Excel formula in memory stream Aspose | save modified workbook back to same stream | reset MemoryStream before saving Aspose.Cells | recalculate formulas after editing workbook in memory
// Developer Intent: Update a formula in an Excel workbook read from a MemoryStream and overwrite the same stream with the modified file.
// Use Cases: Process an uploaded XLSX byte array in a web API, adjust formulas, and return the updated byte array without touching the file system. | Extend a SUM range in an in‑memory workbook, recalculate results, and stream the revised file to another service. | Validate formula changes by reloading the stream and reading the computed value before further processing.
// AI Prompts: Generate C# code using Aspose.Cells to load a workbook from a MemoryStream, modify cell A1's formula to include an extra cell, recalculate, and save back to the same stream. | Explain why resetting the MemoryStream length and position is required before overwriting a workbook with Aspose.Cells. | Provide best‑practice error handling for loading, editing, and saving Excel workbooks in a MemoryStream with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryStreamDemo
{
    // Demonstrates how to load an XLSX workbook from a MemoryStream, change a cell formula, recalculate, clear the stream, and write the updated workbook back to the original MemoryStream using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a workbook with a sample formula.
                Workbook originalWorkbook = new Workbook();
                Worksheet sheet = originalWorkbook.Worksheets[0];
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["A1"].Formula = "=SUM(B1:B2)";

                // 2. Save the workbook to a memory stream.
                using (MemoryStream stream = new MemoryStream())
                {
                    originalWorkbook.Save(stream, SaveFormat.Xlsx);

                    // 3. Reset the stream position to the beginning for reading.
                    stream.Position = 0;

                    // 4. Load the workbook from the same memory stream.
                    Workbook loadedWorkbook = new Workbook(stream);

                    // 5. Modify the existing formula (e.g., extend the range).
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                    loadedSheet.Cells["B3"].PutValue(30); // add a new value to be included
                    loadedSheet.Cells["A1"].Formula = "=SUM(B1:B3)"; // update formula

                    // Optional: recalculate to reflect the new formula result.
                    loadedWorkbook.CalculateFormula();

                    // 6. Prepare the stream for writing the updated workbook.
                    stream.SetLength(0);      // clear previous content
                    stream.Position = 0;      // reset position

                    // 7. Save the modified workbook back to the same memory stream.
                    loadedWorkbook.Save(stream, SaveFormat.Xlsx);

                    // 8. (Optional) Verify by loading again and printing the calculated value.
                    stream.Position = 0;
                    Workbook verifyWorkbook = new Workbook(stream);
                    Console.WriteLine("Updated formula result in A1: " + verifyWorkbook.Worksheets[0].Cells["A1"].Value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
