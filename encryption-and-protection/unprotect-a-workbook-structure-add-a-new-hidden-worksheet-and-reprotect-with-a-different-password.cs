// Title: C# – Unprotect Workbook Structure, Insert Hidden Sheet, and Re‑protect with a New Password using Aspose.Cells
// Description: Demonstrates how to load a password‑protected Excel workbook with Aspose.Cells for .NET, remove its structure protection (Workbook.Unprotect), add a new worksheet, hide it (IsVisible = false), and apply structure protection again with a different password (Workbook.Protect). The modified file is saved to a new location.
// Keywords: Aspose.Cells | C# | .NET | Workbook.Unprotect | Workbook.Protect | Excel structure protection | hidden worksheet | change workbook password | programmatic Excel security | modify protected workbook
// Common Searches: Aspose.Cells unprotect workbook structure C# | add hidden sheet to protected Excel file Aspose.Cells | change workbook protection password .NET | how to re‑protect Excel workbook after adding sheet | C# code to modify protected workbook with Aspose.Cells
// Developer Intent: Remove existing structure protection, add a hidden worksheet, and apply a new structure‑protection password.
// Use Cases: Create a template where confidential sheets stay hidden while the workbook is protected with a custom password. | Update legacy financial reports by inserting a secret worksheet and re‑securing the file with a new password. | Automate migration of protected workbooks to a new security policy without manual Excel interaction.
// AI Prompts: Show C# code using Aspose.Cells to unprotect a workbook's structure, add a hidden worksheet, and protect it again with a different password. | Explain step‑by‑step how to change the structure‑protection password of an Excel file after inserting a hidden sheet in .NET. | Generate a .NET example that loads a protected workbook, removes protection, hides a new sheet, and reapplies protection with a new password.

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates how to load a password‑protected Excel workbook with Aspose.Cells for .NET, remove its structure protection (Workbook.Unprotect), add a new worksheet, hide it (IsVisible = false), and apply structure protection again with a different password (Workbook.Protect). The modified file is saved to a new location.
class WorkbookStructureProtectionDemo
{
    static void Main()
    {
        try
        {
            // Path to the existing workbook that is protected with a password
            string inputPath = "protected_workbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Password that was used to protect the workbook structure
            string oldPassword = "oldPassword";

            // Unprotect the workbook structure using the old password
            workbook.Unprotect(oldPassword);

            // Add a new worksheet to the workbook
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];

            // Hide the newly added worksheet
            newSheet.IsVisible = false;

            // Protect the workbook structure again with a different password
            string newPassword = "newPassword123";
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Path for the modified workbook
            string outputPath = "modified_workbook.xlsx";

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
