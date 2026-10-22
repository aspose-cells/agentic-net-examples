// Title: Remove opening password and workbook structure protection from an Excel workbook in C# using Aspose.Cells and save the unprotected file
// Common Searches: C# Aspose.Cells remove workbook opening password and structure protection in one step | How to load a password‑protected .xlsx and save it without any protection using Aspose.Cells | Aspose.Cells .NET clear workbook password and disable structure protection programmatically | Unprotect an Excel workbook (open password and sheet protection) with Aspose.Cells in C# | Remove both opening password and workbook structure protection from a spreadsheet using Aspose.Cells
// Tags: remove workbook opening password Aspose.Cells | clear workbook structure protection Aspose.Cells | load password protected Excel Aspose.Cells C# | save unprotected workbook Aspose.Cells | Aspose.Cells Settings.Password empty

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the source file, loads the protected workbook with the opening password via LoadOptions, clears the password by setting Settings.Password to an empty string, and saves the workbook as a new .xlsx file without any opening or structure protection.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "protected.xlsx";
        string outputPath = "unprotected.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook using the current opening password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "openPwd" // replace with the actual opening password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove the opening password (if any) by clearing the password setting
            workbook.Settings.Password = string.Empty;

            // Save the workbook without any protection
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
