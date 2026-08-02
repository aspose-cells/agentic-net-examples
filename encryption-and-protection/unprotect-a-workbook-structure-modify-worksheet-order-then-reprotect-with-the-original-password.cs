// Title: C# – Unprotect Excel workbook structure, reorder sheets, and re‑protect with Aspose.Cells
// Description: Loads a password‑protected workbook, removes structure protection, moves a worksheet (e.g., the third sheet) to a new position, reapplies the same protection, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unprotect workbook | C# reorder Excel sheets | protect workbook structure password | Worksheet.MoveTo example | modify sheet order in protected Excel
// Common Searches: how to unprotect and reorder sheets in a protected Excel file using Aspose.Cells | C# move worksheet after unprotecting workbook structure | re‑protect workbook structure after changing sheet order Aspose.Cells
// Developer Intent: Remove structure protection from a workbook, change the order of its worksheets, then restore the original protection with the same password.
// Use Cases: Insert a summary tab at the front of a locked financial report before distribution. | Re‑order generated template sheets after unlocking a workbook for automated processing. | Batch‑process protected workbooks to adjust sheet sequence while preserving existing passwords.
// AI Prompts: Write C# code that uses Aspose.Cells to unprotect a workbook's structure, move the third worksheet to the first position, and re‑protect it with the original password. | Provide a robust example that checks for at least three worksheets, handles missing file errors, logs each step, and saves the reordered workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a password‑protected workbook, removes structure protection, moves a worksheet (e.g., the third sheet) to a new position, reapplies the same protection, and saves the file using Aspose.Cells for .NET.
    public class WorkbookStructureReorderDemo
    {
        public static void Run()
        {
            // Path to the protected workbook
            string inputPath = "ProtectedWorkbook.xlsx";
            // Password used to protect the workbook structure
            string password = "myPassword";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule: load)
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Unprotect the workbook structure using the original password
                    workbook.Unprotect(password);

                    // Example reordering: move the third worksheet (index 2) to the first position (index 0)
                    // Ensure the workbook has at least three sheets
                    if (workbook.Worksheets.Count > 2)
                    {
                        Worksheet sheetToMove = workbook.Worksheets[2];
                        sheetToMove.MoveTo(0); // MoveTo method repositions the sheet
                    }

                    // Re‑protect the workbook structure with the same password
                    workbook.Protect(ProtectionType.Structure, password);

                    // Save the modified workbook (lifecycle rule: save)
                    string outputPath = "ReorderedWorkbook.xlsx";
                    workbook.Save(outputPath);

                    Console.WriteLine($"Workbook structure unprotected, reordered, and re‑protected. Saved as '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookStructureReorderDemo.Run();
        }
    }
}
