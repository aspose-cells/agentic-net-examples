// Title: How to remove workbook and worksheet password protection from an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a password‑protected .xlsx, calls Workbook.Unprotect and Worksheet.Unprotect with empty passwords, and saves the result as a new unprotected workbook. | Show a C# example that iterates through every worksheet in a workbook, clears any sheet‑level protection, and writes the unprotected file to a different location using Aspose.Cells.
// Common Searches: aspnet remove password from excel workbook using Aspose.Cells | c# code to unprotect all sheets in a protected xlsx with Aspose.Cells | how to programmatically delete workbook protection Aspose.Cells .NET | Aspose.Cells unprotect worksheet without knowing password | save unprotected copy of a protected Excel file in C#
// Tags: Aspose.Cells clear workbook protection C# | Aspose.Cells clear worksheet protection C# | Aspose.Cells save unprotected Excel file .NET | Aspose.Cells iterate worksheets unprotect | Aspose.Cells handle protected Excel file

using System;
using System.IO;
using Aspose.Cells;

// // Loads a possibly password‑protected Excel workbook, removes both workbook‑level and worksheet‑level protection using empty passwords, and saves the file as an unprotected copy.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "unprotected.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook that may have password protection
            var workbook = new Workbook(inputPath);

            // Attempt to remove workbook-level protection (no exception if not protected)
            workbook.Unprotect(string.Empty);

            // Remove worksheet-level protection for each sheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Unprotect the worksheet; safe to call even if not protected
                sheet.Unprotect(string.Empty);
            }

            // Save the workbook without any protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
