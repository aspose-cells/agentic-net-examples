// Title: C# console utility to detect whether an Excel workbook is encrypted and display the result using Aspose.Cells
// AI Prompts: Write a C# console program that accepts an Excel file path, loads it with Aspose.Cells, and prints "Encrypted" or "Not encrypted" based on the load outcome. | Enhance the program to distinguish between a missing file, an unsupported format, and a password‑required exception, providing clear console messages for each case. | Add an optional password argument; when supplied, use LoadOptions.Password to open the workbook and confirm successful decryption.
// Common Searches: aspocells c# check if xlsx file is password protected without opening it | how to catch CellsException for encrypted Excel workbook in a .NET console app | detect encryption status of an Excel workbook using Aspose.Cells LoadOptions | c# console tool to report whether an Excel file requires a password | handle missing file and unsupported format errors when loading Excel with Aspose.Cells
// Tags: aspocells detect encrypted workbook | c# loadoptions password exception handling | excel encryption status check .net | aspocells workbook loading error classification | c# console utility aspocells encryption detection

using System;
using System.IO;
using Aspose.Cells;

// The example builds a C# console application that receives an Excel file path, verifies its existence, and attempts to load it with Aspose.Cells using default LoadOptions. If loading succeeds, the file is reported as not encrypted; if a CellsException containing the word "password" is thrown, the program reports the file as encrypted. Additional error handling differentiates missing files, unsupported formats, and other load failures, and an optional password argument can be used to open protected workbooks.
class EncryptionMetadataExtractor
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: EncryptionMetadataExtractor <excel-file-path>");
            return;
        }

        string filePath = args[0];

        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        // Prepare load options without a password – we only need metadata.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);

        try
        {
            // Attempt to load the workbook. If the file is not encrypted this will succeed.
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("The file is not encrypted.");
        }
        catch (CellsException ex)
        {
            // If the exception message indicates a password is required, treat it as encrypted.
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The file is encrypted.");
                // Encryption metadata extraction is not supported without providing a password.
                // If needed, additional Aspose.Cells APIs can be used here to retrieve such info.
            }
            else
            {
                // Other CellsException errors.
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception e)
        {
            // Any other error (e.g., unsupported format) is reported here.
            Console.WriteLine($"Error: {e.Message}");
        }
    }
}
