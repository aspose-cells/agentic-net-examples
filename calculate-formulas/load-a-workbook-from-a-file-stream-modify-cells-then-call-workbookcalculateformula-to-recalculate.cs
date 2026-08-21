// Title: Load a Workbook from a FileStream, Edit Cells, Recalculate Formulas, and Save with Aspose.Cells for .NET
// Description: This C# example demonstrates how to open an XLSX file via a read‑only FileStream, create a Workbook, modify cell values and formulas, invoke Workbook.CalculateFormula to refresh all calculations, display the result, and save the updated workbook to a new file.
// Keywords: Aspose.Cells load workbook from stream | C# Workbook.CalculateFormula | edit cell value Aspose.Cells | set cell formula Aspose.Cells | save workbook after calculation | Aspose.Cells file stream example | .NET Excel formula recalculation
// Common Searches: How to load an Excel file from a FileStream using Aspose.Cells | Aspose.Cells C# recalculate formulas after editing cells | Save workbook after calling Workbook.CalculateFormula | Example of setting a formula programmatically with Aspose.Cells | Manual calculation mode Aspose.Cells .NET
// Developer Intent: Load an Excel workbook from a stream, change cell contents, manually recalculate all formulas, and write the modified file back to disk.
// Use Cases: Process an uploaded Excel file received as a stream in a web API, update input cells, recalculate dependent formulas, and return the revised workbook. | Generate reports from a template stored in memory by inserting parameters, triggering a full formula refresh, and exporting the final document. | Batch‑process multiple workbooks in a background service: stream each file, apply bulk updates, invoke manual calculation for performance, and save the results.
// AI Prompts: Provide a C# snippet that reads an Excel file from a MemoryStream, sets several formulas, calls Workbook.CalculateFormula, and saves the file. | Explain how to disable automatic calculation in Aspose.Cells, perform bulk cell updates, and then manually invoke Workbook.CalculateFormula for optimal speed. | Show how to capture the calculated value of a cell after calling Workbook.CalculateFormula in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaRecalcDemo
{
    // This C# example demonstrates how to open an XLSX file via a read‑only FileStream, create a Workbook, modify cell values and formulas, invoke Workbook.CalculateFormula to refresh all calculations, display the result, and save the updated workbook to a new file.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Open the file as a read‑only stream
            using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook from the stream (uses the Workbook(Stream) constructor)
                Workbook workbook = new Workbook(inputStream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Modify cells: set a value in B1 and a formula in A1 that references B1
                cells["B1"].PutValue(42);
                cells["A1"].Formula = "=B1";

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Optionally display the calculated result
                Console.WriteLine("Calculated value of A1: " + cells["A1"].Value);

                // Save the updated workbook to a new file
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}
