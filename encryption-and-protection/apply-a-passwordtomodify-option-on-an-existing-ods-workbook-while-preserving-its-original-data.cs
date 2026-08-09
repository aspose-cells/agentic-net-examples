// Title: Apply Modify‑Password Write Protection to an Existing ODS Workbook with Aspose.Cells for .NET
// Description: Loads an existing ODS workbook (or creates one if missing), sets a password‑to‑modify via Workbook.Settings.WriteProtection, optionally adds author and read‑only recommendation, and saves the file while preserving all original data.
// Keywords: Aspose.Cells | C# | ODS workbook protection | write protection | modify password | set password to modify ODS | preserve spreadsheet data | load existing ODS | save protected ODS | author attribute write protection
// Common Searches: Aspose.Cells C# add modify password to ODS file | how to protect existing ODS workbook with password to modify | set write protection on ODS spreadsheet using Aspose.Cells | preserve data while applying password protection to ODS in .NET | add author and read‑only recommendation to ODS write protection
// Developer Intent: Add a password‑to‑modify protection to an existing ODS workbook without altering its content.
// Use Cases: Secure an existing ODS spreadsheet by requiring a password before edits can be made. | Generate a new ODS file when the source is unavailable, apply write protection, and distribute it safely. | Include metadata such as author name and a read‑only recommendation alongside the modify password for enhanced protection.
// AI Prompts: Write C# code using Aspose.Cells to open an ODS file, set a modify password, and save the protected workbook. | Show how to add author information and a recommend‑read‑only flag to ODS write protection with Aspose.Cells. | Explain handling of a missing input ODS file by creating a new workbook, applying password protection, and saving it.

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing ODS workbook (or creates one if missing), sets a password‑to‑modify via Workbook.Settings.WriteProtection, optionally adds author and read‑only recommendation, and saves the file while preserving all original data.
class ApplyWriteProtection
{
    static void Main()
    {
        // Path to the existing ODS workbook
        string inputPath = "ExistingWorkbook.ods";

        // Path for the protected output workbook
        string outputPath = "ProtectedWorkbook.ods";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
                Console.WriteLine($"Input file not found. Created a new workbook.");
            }

            // Apply write‑protection (password‑to‑modify)
            workbook.Settings.WriteProtection.Password = "ModifyPassword123";

            // Optional settings (uncomment if needed)
            // workbook.Settings.WriteProtection.Author = "John Doe";
            // workbook.Settings.WriteProtection.RecommendReadOnly = true;

            // Save the workbook with write‑protection applied
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved with write‑protection to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
