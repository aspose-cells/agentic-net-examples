// Title: Unprotect Excel workbook structure, rename a worksheet, and re‑apply structure protection with a new password in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells unprotect workbook structure, rename worksheet, and protect with new password | How to change a sheet name in a password‑protected Excel file with Aspose.Cells for .NET | Replace old workbook protection password and rename first sheet using Aspose.Cells | Aspose.Cells protect workbook structure after modifying sheet name in C# | Load Excel file, remove structure protection, rename sheet, and save with new password using Aspose.Cells
// Tags: unprotect workbook structure Aspose.Cells | rename worksheet Aspose.Cells C# | protect workbook structure with new password Aspose.Cells | modify protected Excel file using Aspose.Cells | change sheet name in password‑protected workbook C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, removes its structure protection with the old password, renames the first worksheet, then re‑applies structure protection using a new password, ensures the output directory exists, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook structure using the existing password
            string oldPassword = "oldPass";
            workbook.Unprotect(oldPassword); // Removes all protections, including structure

            // Rename the first worksheet (change index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "RenamedSheet";

            // Re‑apply structure protection with a new password
            string newPassword = "newPass";
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
