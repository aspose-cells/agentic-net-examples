// Title: Modify an Excel formula in a MemoryStream and save back with Aspose.Cells for .NET
// Description: Creates a workbook in memory, writes it to a MemoryStream (XLSX), reloads it, replaces the formula in C1 using SetFormula, clears the stream, saves the updated workbook to the same stream, and verifies the change by reading and calculating the new value—all without touching the file system.
// Keywords: Aspose.Cells MemoryStream formula | C# update Excel cell formula in memory | SetFormula Aspose.Cells example | load workbook from stream Aspose | save workbook to same stream .NET | in‑memory Excel manipulation | stream.Position reset Aspose.Cells
// Common Searches: change Excel formula from MemoryStream Aspose.Cells | write modified workbook back to original stream C# | replace cell formula without saving to disk | Aspose.Cells SetFormula usage | reset MemoryStream before overwriting workbook
// Developer Intent: Load an Excel file from a MemoryStream, replace an existing formula, and write the updated workbook back into the same stream using Aspose.Cells for .NET.
// Use Cases: Edit formulas of Excel files received via web APIs before sending them back to the client. | Perform server‑side calculations on uploaded spreadsheets without creating temporary files. | Batch‑process workbooks stored as BLOBs in a database, updating formulas directly in their memory streams.
// AI Prompts: Generate C# code that loads an XLSX from a MemoryStream, changes cell C1 formula to "=A1*B1" with Aspose.Cells, and writes the workbook back to the same stream. | Show best‑practice error handling for modifying formulas in a workbook loaded from a MemoryStream using Aspose.Cells. | Explain how to correctly reset a MemoryStream's length and position when overwriting a workbook after a formula change.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Creates a workbook in memory, writes it to a MemoryStream (XLSX), reloads it, replaces the formula in C1 using SetFormula, clears the stream, saves the updated workbook to the same stream, and verifies the change by reading and calculating the new value—all without touching the file system.
    public class MemoryStreamFormulaExample
    {
        public static void Run()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a sample workbook with a simple formula.
                // ------------------------------------------------------------
                Workbook originalWorkbook = new Workbook();
                Worksheet sheet = originalWorkbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["B1"].PutValue(20);
                // Initial formula in C1: =A1+B1
                sheet.Cells["C1"].Formula = "=A1+B1";

                // ------------------------------------------------------------
                // 2. Save the workbook to a memory stream (XLSX format).
                // ------------------------------------------------------------
                using (MemoryStream stream = new MemoryStream())
                {
                    originalWorkbook.Save(stream, SaveFormat.Xlsx);

                    // Reset the stream position to the beginning for reading.
                    stream.Position = 0;

                    // ------------------------------------------------------------
                    // 3. Load the workbook from the same memory stream.
                    // ------------------------------------------------------------
                    Workbook loadedWorkbook = new Workbook(stream);

                    // ------------------------------------------------------------
                    // 4. Modify the existing formula (e.g., change to =A1*B1).
                    // ------------------------------------------------------------
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                    // Use SetFormula to replace the formula.
                    loadedSheet.Cells["C1"].SetFormula("=A1*B1", null);

                    // ------------------------------------------------------------
                    // 5. Save the modified workbook back into the same memory stream.
                    // ------------------------------------------------------------
                    // Clear the previous content.
                    stream.SetLength(0);
                    // Reset position before writing.
                    stream.Position = 0;
                    loadedWorkbook.Save(stream, SaveFormat.Xlsx);
                    // Reset position for further reading.
                    stream.Position = 0;

                    // ------------------------------------------------------------
                    // 6. Demonstrate that the formula has been updated.
                    // ------------------------------------------------------------
                    Workbook verifyWorkbook = new Workbook(stream);
                    string updatedFormula = verifyWorkbook.Worksheets[0].Cells["C1"].Formula;
                    Console.WriteLine("Updated formula in C1: " + updatedFormula);
                    // Calculate to see the new result.
                    verifyWorkbook.CalculateFormula();
                    Console.WriteLine("Calculated value in C1: " + verifyWorkbook.Worksheets[0].Cells["C1"].Value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MemoryStreamFormulaExample.Run();
        }
    }
}
