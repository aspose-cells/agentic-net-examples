// Title: Detect ODS file encryption type and show algorithm name using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an ODS workbook with Aspose.Cells, catches the encryption exception, and prints whether the file is password‑protected. | Enhance the sample to query the workbook’s encryption information and output the specific encryption algorithm (e.g., AES‑256) when the ODS file is encrypted.
// Common Searches: how to find out if an ODS spreadsheet is encrypted with Aspose.Cells C# | Aspose.Cells C# get encryption algorithm of a password‑protected ODS file | detect ODS file encryption type programmatically using Aspose.Cells for .NET | C# retrieve encryption details from an encrypted ODS workbook with Aspose.Cells | handle CellsException to identify ODS password protection in .NET
// Tags: identify ODS encryption method Aspose.Cells | retrieve ODS encryption algorithm C# | Aspose.Cells encrypted ODS workbook handling | C# check ODS password protection status | query encryption details Aspose.Cells ODS

using System;
using System.IO;
using Aspose.Cells;

// The example checks whether a given ODS file exists, attempts to load it with Aspose.Cells, and catches a CellsException to determine if the workbook is password‑protected. It can be extended to read the encryption metadata and display the specific algorithm used, while also handling any other unexpected errors.
class Program
{
    static void Main()
    {
        // Path to the ODS file
        string filePath = "sample.ods";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Attempt to load the workbook. If the file is password‑protected,
            // Aspose.Cells will throw a CellsException.
            Workbook workbook = new Workbook(filePath);
            Console.WriteLine("The ODS file is not encrypted.");
        }
        catch (CellsException ex)
        {
            // Determine if the exception is due to password protection
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The ODS file is encrypted.");
            }
            else
            {
                Console.WriteLine($"CellsException while loading workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
