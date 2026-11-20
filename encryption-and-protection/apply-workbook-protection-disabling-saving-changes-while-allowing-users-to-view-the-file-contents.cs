// Title: How to protect an Excel workbook from edits and saving while allowing read‑only view using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect workbook structure and prevent users from saving changes | Aspose.Cells write protection to make Excel file read‑only without allowing edits | How to apply full workbook protection with password in .NET using Aspose.Cells | Prevent modifications to an existing .xlsx and keep it view‑only with Aspose.Cells C# | Save a protected copy of an Excel file while blocking further saves using Aspose.Cells
// Tags: Aspose.Cells workbook.Protect with ProtectionType.All | read‑only Excel workbook protection .NET | prevent Excel file edits using Aspose.Cells C# | password‑protected workbook structure Aspose.Cells | save protected copy of .xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input .xlsx file, loads it into an Aspose.Cells Workbook, applies full write protection (structure and windows) with a password via workbook.Protect(ProtectionType.All, ...), and saves the protected workbook, ensuring the file can be opened for viewing but cannot be edited or saved with changes.
class WorkbookProtectionExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "protected_output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Apply write protection (protects the workbook structure and windows)
            workbook.Protect(ProtectionType.All, "StrongPassword123");

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Protected workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
