// Title: Validate extracted Excel file extension against detected format with Aspose.Cells (.NET)
// Description: A C# sample that uses Aspose.Cells FileFormatUtil to read an extracted spreadsheet via a stream, detect its true format, map the detected format and the file's extension to SaveFormat enums, and confirm they are identical. The console output shows the extension, detected format type, mapped SaveFormat values, and the integrity‑verification result.
// Keywords: Aspose.Cells | FileFormatUtil | DetectFileFormat | ExtensionToSaveFormat | SaveFormat | C# | .NET | Excel file validation | spreadsheet integrity check | file extension verification | extracted file stream | zip extraction | format detection
// Common Searches: Aspose.Cells verify Excel file extension matches content | C# detect spreadsheet format from stream | compare detected format with file extension Aspose.Cells | validate extracted .xlsx file integrity | how to check file extension against actual format .NET
// Developer Intent: Ensure an extracted spreadsheet’s extension accurately reflects its underlying format to avoid processing errors.
// Use Cases: Validate files extracted from ZIP archives before loading them into a workbook. | Reject uploaded spreadsheets in a web API when the extension does not match the detected content. | Automate integrity checks in a batch conversion pipeline that extracts and re‑saves spreadsheets in various formats.
// AI Prompts: Write a C# helper that throws an exception if FileFormatUtil detects a format different from the file’s extension. | Create a unit test covering .xlsx, .xls, and .csv files extracted from a zip archive, verifying the integrity check passes or fails as expected. | Generate logging code that records the detected FileFormatType, the mapped SaveFormat, and the original extension when verification fails.

using System;
using System.IO;
using Aspose.Cells;

// A C# sample that uses Aspose.Cells FileFormatUtil to read an extracted spreadsheet via a stream, detect its true format, map the detected format and the file's extension to SaveFormat enums, and confirm they are identical. The console output shows the extension, detected format type, mapped SaveFormat values, and the integrity‑verification result.
class VerifyFileExtension
{
    static void Main()
    {
        // Path to the original file (used only for reference)
        string originalFilePath = "original.xlsx";

        // Path to the file that was extracted (could be from a zip, stream, etc.)
        string extractedFilePath = "extracted_file";

        // Verify that the extracted file exists before processing
        if (!File.Exists(extractedFilePath))
        {
            Console.WriteLine($"Error: Extracted file not found at path '{extractedFilePath}'.");
            return;
        }

        try
        {
            // Detect the format of the extracted file using a stream
            using (FileStream stream = File.OpenRead(extractedFilePath))
            {
                // Get format information from the file content
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream);

                // Convert the detected FileFormatType to a SaveFormat enum value
                SaveFormat detectedSaveFormat = FileFormatUtil.FileFormatToSaveFormat(formatInfo.FileFormatType);

                // Obtain the file extension of the extracted file (including the leading dot)
                string fileExtension = Path.GetExtension(extractedFilePath);

                // Convert the extension to a SaveFormat enum value
                SaveFormat extensionSaveFormat = FileFormatUtil.ExtensionToSaveFormat(fileExtension);

                // Verify that the detected format matches the extension and that the extension is recognized
                bool isIntegrityIntact = detectedSaveFormat == extensionSaveFormat && extensionSaveFormat != SaveFormat.Unknown;

                // Output the verification result
                Console.WriteLine($"Extracted file extension: {fileExtension}");
                Console.WriteLine($"Detected file format type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Detected SaveFormat: {detectedSaveFormat}");
                Console.WriteLine($"Extension mapped SaveFormat: {extensionSaveFormat}");
                Console.WriteLine($"Integrity verification passed: {isIntegrityIntact}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
