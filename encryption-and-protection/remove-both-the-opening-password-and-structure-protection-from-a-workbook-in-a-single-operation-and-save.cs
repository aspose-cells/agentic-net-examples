// Title: Remove opening password and structure protection from an Excel workbook in a single step using Aspose.Cells for .NET
// Common Searches: how to programmatically remove opening password from a protected XLSX with Aspose.Cells C# | Aspose.Cells C# load password‑protected workbook and save without any protection | remove Excel workbook structure protection and password in one operation using Aspose.Cells | C# code to unprotect a password‑protected Excel file with Aspose.Cells | Aspose.Cells .NET clear workbook protection and save as unprotected file
// Tags: Aspose.Cells remove workbook password | Aspose.Cells clear Excel structure protection | Aspose.Cells load protected XLSX | Aspose.Cells save unprotected workbook | Aspose.Cells single‑step decryption .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a password‑protected XLSX file using LoadOptions, clears the opening password by setting workbook.Settings.Password to an empty string, and then saves the workbook as an unprotected file in one operation.
class Program
{
    static void Main()
    {
        // Paths and password
        string inputPath = "protected.xlsx";
        string outputPath = "unprotected.xlsx";
        string openingPassword = "myPassword";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with the opening password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = openingPassword
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Remove the opening password
            workbook.Settings.Password = string.Empty;

            // Save the workbook without any protection
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
