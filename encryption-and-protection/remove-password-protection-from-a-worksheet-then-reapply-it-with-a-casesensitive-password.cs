// Title: Remove worksheet protection and reapply a case‑sensitive password with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unprotect worksheet then protect with custom password | remove existing sheet protection and set case‑sensitive password in Excel using Aspose.Cells | how to use Worksheet.Unprotect and Worksheet.Protect with ProtectionType.All in C# | Aspose.Cells protect entire worksheet with password and save workbook | handle missing Excel file when applying protection with Aspose.Cells .NET
// Tags: Worksheet.Unprotect Aspose.Cells | Worksheet.Protect with ProtectionType.All C# | case-sensitive password Excel sheet Aspose.Cells | save protected workbook Aspose.Cells .NET | file existence check before loading workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook, removes any current protection from the first worksheet using Unprotect(), then applies full sheet protection with a case‑sensitive password via Protect(ProtectionType.All, password). It verifies the input file, creates the output directory if needed, and saves the modified workbook, handling file‑not‑found and other exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "MySecretPassword";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            var worksheet = workbook.Worksheets[0];

            // Remove any existing protection
            worksheet.Unprotect();

            // Apply protection to the entire sheet with the specified password
            // The third parameter is the old password (not required here, so pass null)
            worksheet.Protect(ProtectionType.All, password, null);

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
