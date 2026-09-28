// Title: Remove worksheet protection and reapply a case‑sensitive password with Aspose.Cells for .NET
// AI Prompts: Unprotect the first worksheet of an Excel workbook and then protect it again using a case‑sensitive password with Aspose.Cells in C#. | Generate C# code that loads a workbook, clears any existing worksheet protection, applies a new password, and saves the file using Aspose.Cells.
// Common Searches: Aspose.Cells change worksheet password case sensitive C# example | how to remove Excel sheet protection and set a new password with Aspose.Cells | protect worksheet with password programmatically Aspose.Cells .NET | unprotect and re‑protect Excel worksheet using Aspose.Cells API
// Tags: worksheet.Unprotect method Aspose.Cells example | worksheet.Protect API usage Aspose.Cells C# | apply case‑sensitive password to worksheet Aspose.Cells | save workbook after updating protection .NET | load workbook and modify worksheet security Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads input.xlsx, removes any existing protection from the first worksheet, then protects it with the case‑sensitive password "MyCaseSensitivePwd" using worksheet.Protect, and finally saves the updated workbook to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "MyCaseSensitivePwd";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Remove any existing protection
            worksheet.Unprotect();

            // Apply protection with a password (case‑sensitive by default)
            // The third parameter is the old password; pass empty string if none
            worksheet.Protect(ProtectionType.All, password, string.Empty);

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Output error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
