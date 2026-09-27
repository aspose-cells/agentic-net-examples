// Title: Remove workbook structure and windows protection from an Excel .xlsx file using Aspose.Cells for .NET and verify sheet reordering
// AI Prompts: Call Workbook.Unprotect with the correct password to clear both structure and windows protection from a protected .xlsx workbook using Aspose.Cells, then save the result. | After unprotecting, use Worksheet.MoveTo to reposition the first worksheet and confirm that the workbook is no longer locked.
// Common Searches: asp.net unprotect workbook structure and windows Aspose.Cells password | c# verify worksheet can be moved after removing Excel workbook protection with Aspose.Cells | how to disable workbook protection programmatically in .xlsx using Aspose.Cells | example of moving a worksheet after unprotecting an Excel file in C# | Aspose.Cells remove workbook protection and test sheet reordering
// Tags: Aspose.Cells Workbook.Unprotect password | remove Excel workbook structure protection C# | unprotect workbook windows Aspose.Cells | worksheet MoveTo after unprotecting workbook | verify sheet reordering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected XLSX file, calls Workbook.Unprotect with the supplied password to clear structure and windows protection, moves the first worksheet to a new position to confirm the protection is removed, and saves the unprotected workbook.
class RemoveWorkbookStructureProtection
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "ProtectedWorkbook.xlsx";
        string outputPath = "UnprotectedWorkbook.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Password used to protect the workbook structure
            string password = "myPassword";

            // Remove workbook structure protection (both structure and windows)
            workbook.Unprotect(password);

            // Verify that sheets are now movable by moving the first worksheet to the second position
            try
            {
                if (workbook.Worksheets.Count > 1)
                {
                    // Move worksheet at index 0 to index 1
                    workbook.Worksheets[0].MoveTo(1);
                    Console.WriteLine("Worksheet moved successfully. Structure protection removed.");
                }
                else
                {
                    Console.WriteLine("Workbook does not contain enough worksheets to test moving.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to move worksheet: " + ex.Message);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbook: " + ex.Message);
        }
    }
}
