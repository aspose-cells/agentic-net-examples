// Title: Apply full password‑protected worksheet protection to an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect worksheet with password and save workbook | Aspose.Cells protect first worksheet all options in .NET | How to use ProtectionType.All to lock an Excel worksheet with Aspose.Cells | Saving a password‑protected Excel file after applying worksheet protection in C# | Aspose.Cells ensure output folder exists before saving protected workbook
// Tags: Aspose.Cells worksheet full protection C# | protect Excel worksheet with password Aspose.Cells | ProtectionType.All usage Aspose.Cells | save protected workbook Aspose.Cells .NET | ensure output directory before workbook save C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsProtectionExample
{
    // The program loads an existing Excel workbook, retrieves the first worksheet, applies full password‑protected worksheet protection via worksheet.Protect(ProtectionType.All, password, string.Empty), creates the output directory if it does not exist, and saves the protected workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input, output paths and password (replace with actual values or pass via args)
            string inputFilePath = "{InputFilePath}";
            string outputFilePath = "{OutputFilePath}";
            string password = "{Password}";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFilePath))
                    throw new FileNotFoundException($"Input file not found: {inputFilePath}");

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFilePath);

                // Get the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Protect the worksheet with the given password (oldPassword not required, pass empty string)
                worksheet.Protect(ProtectionType.All, password, string.Empty);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the output path
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
