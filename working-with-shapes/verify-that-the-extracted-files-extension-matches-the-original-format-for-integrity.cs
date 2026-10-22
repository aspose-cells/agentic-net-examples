// Title: Check that an extracted Excel workbook’s file extension matches its original format using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an extracted workbook with Aspose.Cells, reads Workbook.FileFormat, converts it to a file extension, and compares it to a supplied original extension. | Extend the GetExtensionFromFileFormat helper to handle additional FileFormatType values such as Xls, Xlsm, and Ods, and provide a default for unknown types. | Create a reusable method that accepts a file path and an expected extension, validates the workbook’s detected format with Aspose.Cells, and throws an error when they differ.
// Common Searches: how to compare detected Excel file format with original extension using Aspose.Cells C# | Aspose.Cells verify integrity of extracted workbook by checking file extension | C# get file extension from Aspose.Cells FileFormatType enum | detect workbook format automatically with Aspose.Cells and validate against expected type | validate extracted spreadsheet file type matches original after extraction
// Tags: verify extracted workbook extension Aspose.Cells | retrieve workbook FileFormat C# Aspose.Cells | FileFormatType extension conversion .NET | match expected spreadsheet extension with detected | validate Excel extraction integrity Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads the extracted Excel file, uses Aspose.Cells to automatically detect its format via Workbook.FileFormat, maps the detected FileFormatType to a file extension, and checks whether this extension matches the original one, reporting a match or mismatch.
class Program
{
    static void Main()
    {
        // Path to the extracted file
        string extractedFilePath = @"C:\Temp\extracted.xlsx";

        // Ensure the file exists before attempting to load it
        if (!File.Exists(extractedFilePath))
        {
            Console.WriteLine($"File not found: {extractedFilePath}");
            return;
        }

        // Original file extension (could be retrieved from metadata or original file name)
        string originalExtension = ".xlsx";

        try
        {
            // Load the workbook; Aspose.Cells automatically detects the format
            Workbook workbook = new Workbook(extractedFilePath);

            // Get the detected file format from the loaded workbook
            FileFormatType detectedFormat = workbook.FileFormat;

            // Convert the detected format to a file extension string
            string detectedExtension = GetExtensionFromFileFormat(detectedFormat);

            // Verify that the detected extension matches the original extension
            if (string.Equals(detectedExtension, originalExtension, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("File extension matches the original format.");
            }
            else
            {
                Console.WriteLine($"Extension mismatch: original {originalExtension}, detected {detectedExtension}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., corrupted file, unsupported format)
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    // Helper method to map Aspose.Cells FileFormatType to a file extension
    static string GetExtensionFromFileFormat(FileFormatType format)
    {
        switch (format)
        {
            case FileFormatType.Xlsx: return ".xlsx";
            case FileFormatType.Xlsb: return ".xlsb";
            case FileFormatType.Xlsm: return ".xlsm";
            case FileFormatType.Csv:  return ".csv";
            case FileFormatType.Html: return ".html";
            case FileFormatType.Pdf:  return ".pdf";
            case FileFormatType.Ods:  return ".ods";
            default:
                // For older Excel formats (XLS) or any unrecognized format, default to .xls
                return ".xls";
        }
    }
}
