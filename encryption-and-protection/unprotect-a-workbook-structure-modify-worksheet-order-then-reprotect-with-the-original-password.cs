// Title: C# – Unprotect Excel workbook structure, reorder sheets, and re‑protect with Aspose.Cells
// Description: Loads a password‑protected Excel file, removes structure protection, moves the first worksheet to the last position, reapplies the same password protection, and saves the result. Includes handling for missing or incorrect passwords.
// Keywords: Aspose.Cells C# unprotect workbook | Excel workbook structure protection | move worksheet programmatically | Workbook.Unprotect password | Workbook.Protect structure | reorder sheets Aspose.Cells | handle invalid password Aspose.Cells | C# Excel sheet ordering example
// Common Searches: how to unprotect workbook structure with Aspose.Cells C# | move first worksheet to last using Aspose.Cells | re‑protect Excel workbook after changing sheet order | Aspose.Cells catch invalid password exception | C# example for workbook.Unprotect and Protect
// Developer Intent: Remove structure protection, change worksheet order, then restore protection with the original password in a .NET application.
// Use Cases: Standardize sheet sequence in protected reports while keeping the original password. | Batch‑process protected workbooks to enforce a specific tab order without exposing passwords. | Gracefully continue processing when a workbook is unprotected or the password is wrong, ensuring the file is still saved.
// AI Prompts: Generate C# code that catches a CellsException for an invalid password when calling Workbook.Unprotect in Aspose.Cells. | Show how to move a worksheet to a new index and then protect the workbook structure with a password using Aspose.Cells. | Explain how to check if a workbook is structure‑protected before attempting to unprotect it with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookStructureReorder
{
    // Loads a password‑protected Excel file, removes structure protection, moves the first worksheet to the last position, reapplies the same password protection, and saves the result. Includes handling for missing or incorrect passwords.
    class Program
    {
        static void Main()
        {
            // Path to the protected workbook
            string inputPath = "protected_workbook.xlsx";

            // Path where the modified workbook will be saved
            string outputPath = "reordered_workbook.xlsx";

            // Password used to protect the workbook structure
            string password = "mySecretPassword";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Attempt to unprotect the workbook structure; ignore if password is invalid or not protected
                try
                {
                    workbook.Unprotect(password);
                }
                catch (CellsException ex)
                {
                    // Aspose.Cells throws CellsException with a message indicating an invalid password
                    if (ex.Message != null && ex.Message.IndexOf("Invalid password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine("Warning: Invalid password or workbook is not protected. Continuing without unprotecting.");
                    }
                    else
                    {
                        // Re‑throw unexpected exceptions
                        throw;
                    }
                }

                // Example: move the first worksheet to the last position
                if (workbook.Worksheets.Count > 1)
                {
                    Worksheet firstSheet = workbook.Worksheets[0];
                    firstSheet.MoveTo(workbook.Worksheets.Count - 1);
                }

                // Re‑protect the workbook structure with the same password
                workbook.Protect(ProtectionType.Structure, password);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
