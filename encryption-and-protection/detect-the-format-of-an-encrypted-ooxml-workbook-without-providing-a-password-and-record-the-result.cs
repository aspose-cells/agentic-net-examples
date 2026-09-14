// Title: Detect the format of a password‑protected OOXML workbook without providing a password using Aspose.Cells for .NET
// AI Prompts: Load an encrypted .xlsx file with LoadOptions set to Unknown and read the Workbook.FileFormat property. | Write C# code that opens a password‑protected Excel workbook without a password and outputs the detected file type. | Implement error handling for loading a protected OOXML workbook while still obtaining its format information.
// Common Searches: Aspose.Cells detect workbook type when password is unknown | load encrypted Excel file without password and get format .NET | how to read file format of password‑protected .xlsx using Aspose | determine OOXML workbook format without supplying password in C#
// Tags: loadoptions unknown format Aspose.Cells | workbook.fileformat property usage | password‑protected OOXML detection .NET | encrypted Excel workbook format identification | handling unknown password Excel load error

using System;
using System.IO;
using Aspose.Cells;

// The program checks for an encrypted .xlsx file, loads it with LoadOptions set to Unknown (no password supplied), retrieves the Workbook.FileFormat property to identify the workbook’s format, prints the result, and gracefully handles any loading exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted OOXML workbook
        string filePath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Create LoadOptions with unknown format; do not provide a password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Unknown);

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Retrieve the detected file format
            FileFormatType detectedFormat = workbook.FileFormat;

            // Output the result
            Console.WriteLine($"Detected format: {detectedFormat}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., invalid password, corrupted file)
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
