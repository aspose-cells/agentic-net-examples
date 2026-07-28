// Title: Remove Opening Password & Structure Protection from an Excel Workbook with Aspose.Cells for .NET
// Description: Loads a password‑protected .xlsx via LoadOptions, calls Workbook.Unprotect to clear any workbook‑structure lock, clears the file‑open encryption password through Workbook.Settings.Password, and saves the result as an unprotected workbook—all in a single operation.
// Keywords: Aspose.Cells remove password | C# unprotect Excel workbook | clear workbook structure protection | delete opening encryption password | save unprotected Excel file .NET | Workbook.Unprotect | LoadOptions.Password | Excel security removal
// Common Searches: how to remove opening password from Excel using Aspose.Cells | C# Aspose.Cells unprotect workbook structure | remove both password and structure protection in one step | Aspose.Cells load encrypted workbook and save without password
// Developer Intent: Strip both the file‑open password and any workbook structure lock from an Excel file and write it back without protection.
// Use Cases: Batch conversion of incoming password‑protected spreadsheets into editable files for downstream processing. | Automating protection removal before applying data extraction, validation, or transformation logic. | Preparing a workbook for public distribution so users can open it without entering a password.
// AI Prompts: Generate C# code using Aspose.Cells that opens a password‑protected .xlsx, removes the opening password and workbook structure protection, and saves the file unprotected. | Explain how Workbook.Unprotect and Workbook.Settings.Password work together to clear protection in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected .xlsx via LoadOptions, calls Workbook.Unprotect to clear any workbook‑structure lock, clears the file‑open encryption password through Workbook.Settings.Password, and saves the result as an unprotected workbook—all in a single operation.
class RemoveWorkbookProtection
{
    static void Main()
    {
        // Path to the workbook that is protected with an opening password and possibly structure protection
        string inputPath = "protected_workbook.xlsx";
        string outputPath = "unprotected_workbook.xlsx";

        // Password used for opening the file and for workbook structure protection (if any)
        string password = "myPassword";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook using the opening password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Attempt to remove workbook structure protection (if it is enabled)
            try
            {
                workbook.Unprotect(password);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Failed to unprotect workbook structure: {ex.Message}");
            }

            // Remove the opening (encryption) password
            if (!string.IsNullOrEmpty(workbook.Settings.Password))
            {
                workbook.Settings.Password = null;
            }

            // Save the unprotected workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without protection: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
