// Title: Open a password‑protected XLSX with Aspose.Cells for .NET and save it without a password
// Common Searches: Aspose.Cells C# load an encrypted XLSX using a password and save it unprotected | remove opening password from Excel workbook programmatically with Aspose.Cells .NET | C# example to open password protected workbook and export without password using Aspose.Cells | how to use LoadOptions.Password in Aspose.Cells to decrypt and save an XLSX file | unprotect Excel file with Aspose.Cells while preserving original content
// Tags: Aspose.Cells loadoptions password | remove workbook opening password Aspose.Cells | save unprotected XLSX Aspose.Cells | C# decrypt Excel file Aspose.Cells | open encrypted XLSX .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the source file, creates LoadOptions with the opening password, loads the protected workbook, and then saves it to a new file without specifying any password, effectively removing the opening protection while keeping the workbook unchanged.
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

            // Load the workbook using the opening password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "myPassword" // opening password
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook without any password protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
