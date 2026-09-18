// Title: Validate that an extracted Excel workbook’s extension matches its original Aspose.Cells format using C#
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, retrieves its FileFormatType, converts it to the expected extension, and checks it against the file’s actual extension. | Implement a C# helper that maps Aspose.Cells FileFormatType enums to standard file extensions and uses it to perform an integrity verification on an extracted spreadsheet.
// Common Searches: C# Aspose.Cells compare workbook FileFormatType to file extension for validation | how to ensure extracted Excel file has correct extension after using Aspose.Cells | map Aspose.Cells file format enum to extension example in .NET | verify spreadsheet integrity by checking original format versus file extension in C#
// Tags: Aspose.Cells verify extension matches original format | C# translate FileFormatType enum into file extension | Excel workbook integrity validation using Aspose | check extracted spreadsheet extension C# | detect original workbook format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an extracted workbook, obtains its original FileFormatType via Aspose.Cells, translates that format to the expected file extension, and compares it with the actual file extension to confirm integrity, outputting a pass or fail message.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the extracted file
            string extractedFilePath = @"C:\Temp\extracted.xlsx";

            // Ensure the file exists before loading
            if (!File.Exists(extractedFilePath))
            {
                Console.WriteLine($"Error: File not found at '{extractedFilePath}'.");
                return;
            }

            // Get the file extension of the extracted file
            string extractedExtension = Path.GetExtension(extractedFilePath).ToLowerInvariant();

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(extractedFilePath);

            // Determine the original format of the workbook
            FileFormatType originalFormat = workbook.FileFormat;

            // Map the original format to its typical file extension
            string originalExtension = GetExtensionForFormat(originalFormat);

            // Compare extensions
            if (extractedExtension == originalExtension)
            {
                Console.WriteLine("Integrity check passed: file extension matches original format.");
            }
            else
            {
                Console.WriteLine($"Integrity check failed: extracted file extension '{extractedExtension}' does not match original format '{originalExtension}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Helper method to map Aspose.Cells FileFormatType to file extension
    private static string GetExtensionForFormat(FileFormatType format)
    {
        string name = format.ToString().ToLowerInvariant();

        switch (name)
        {
            case "xlsx":
                return ".xlsx";
            case "xls":
            case "xls2003":
                return ".xls";
            case "csv":
                return ".csv";
            case "html":
                return ".html";
            case "pdf":
                return ".pdf";
            case "xlsb":
                return ".xlsb";
            case "xltx":
                return ".xltx";
            case "xlt":
                return ".xlt";
            case "tsv":
                return ".tsv";
            case "ods":
                return ".ods";
            default:
                return string.Empty;
        }
    }
}
