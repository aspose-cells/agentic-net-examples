// Title: C# – Unprotect Excel workbook structure, add hidden sheet, and re‑protect with new password using Aspose.Cells
// Description: Shows how to load a password‑protected workbook with Aspose.Cells for .NET, remove its structure protection, insert an invisible worksheet, and apply a new structure‑protection password before saving the file.
// Keywords: Aspose.Cells unprotect workbook | C# Excel structure protection | add hidden worksheet Aspose.Cells | protect workbook with password .NET | modify protected Excel file C# | Aspose.Cells workbook protection example
// Common Searches: remove workbook structure protection Aspose.Cells C# | add hidden sheet to protected Excel workbook .NET | re‑protect Excel file with new password using Aspose.Cells | change Excel workbook protection password programmatically | Aspose.Cells hide worksheet after unprotecting
// Developer Intent: The developer needs to lift existing structure protection, insert a hidden worksheet, and then re‑apply structure protection with a different password.
// Use Cases: Prepare a template workbook, unprotect it, embed a hidden configuration sheet, and re‑protect before distribution. | Update a secured financial report by adding a concealed audit sheet and resetting the protection password. | Batch‑process protected workbooks to inject hidden metadata sheets and enforce a uniform password.
// AI Prompts: Write C# code with Aspose.Cells that removes workbook structure protection, adds a hidden worksheet, and protects the workbook with a new password. | Explain how to catch and handle CellsException when trying to unprotect an Excel workbook without providing a password. | Provide step‑by‑step instructions for changing the structure‑protection password of an Excel file while preserving hidden sheets using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load a password‑protected workbook with Aspose.Cells for .NET, remove its structure protection, insert an invisible worksheet, and apply a new structure‑protection password before saving the file.
class WorkbookStructureProtectionDemo
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "ProtectedWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to unprotect the workbook structure.
            // Passing an empty string tries to remove protection without a password.
            try
            {
                workbook.Unprotect(string.Empty);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Unable to unprotect without password: {ex.Message}");
                // If you know the correct password, uncomment and use the line below:
                // workbook.Unprotect("oldPassword123");
            }

            // Add a new hidden worksheet
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.IsVisible = false; // hide the worksheet

            // Protect the workbook structure with a new password
            string newPassword = "newPassword456";
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Save the modified workbook
            string outputPath = "ModifiedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
