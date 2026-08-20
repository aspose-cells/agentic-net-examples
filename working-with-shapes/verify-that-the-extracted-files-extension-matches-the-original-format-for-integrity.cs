// Title: C# – Verify an extracted spreadsheet’s extension matches its detected format using Aspose.Cells
// Description: Shows how to employ Aspose.Cells FileFormatUtil to identify the real format of an extracted workbook, convert the format to a SaveFormat, obtain the standard file extension, compare it with the file’s current extension, and output the verification result while handling missing files and exceptions.
// Keywords: Aspose.Cells | FileFormatUtil | detect spreadsheet format | verify file extension C# | SaveFormatToExtension | Excel file validation .NET | extension mismatch detection | spreadsheet integrity check | C# file format detection
// Common Searches: Aspose.Cells detect file format C# example | C# check if Excel file extension matches its content | How to verify spreadsheet extension with Aspose.Cells | FileFormatUtil usage for extension validation | Validate extracted Excel file extension in .NET
// Developer Intent: The developer wants to confirm that the extension of an extracted workbook accurately reflects the format detected by Aspose.Cells before any further processing.
// Use Cases: Screen extracted files from zip archives to ensure they carry the correct extension before conversion. | Reject user‑uploaded spreadsheets whose extensions do not correspond to the actual file format in a web portal. | Automate batch integrity checks in a data‑pipeline that extracts and saves workbooks in multiple formats.
// AI Prompts: Write C# code that logs a warning instead of printing to console when the detected extension differs from the file name, using Aspose.Cells. | Create a method that throws a custom ExtensionMismatchException when the extension check fails. | Show how to embed this verification step into an ASP.NET Core file‑upload controller.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to employ Aspose.Cells FileFormatUtil to identify the real format of an extracted workbook, convert the format to a SaveFormat, obtain the standard file extension, compare it with the file’s current extension, and output the verification result while handling missing files and exceptions.
    public class FileExtensionIntegrityVerifier
    {
        /// <param name="originalFilePath">Path to the original file (used only for reference).</param>
        /// <param name="extractedFilePath">Path to the extracted file whose integrity is to be verified.</param>
        public static void VerifyFileExtension(string originalFilePath, string extractedFilePath)
        {
            // Ensure the extracted file exists
            if (!File.Exists(extractedFilePath))
            {
                Console.WriteLine($"Extracted file not found: {extractedFilePath}");
                return;
            }

            try
            {
                // Detect the actual format of the extracted file
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(extractedFilePath);

                // Convert the detected FileFormatType to a SaveFormat enum
                SaveFormat detectedSaveFormat = FileFormatUtil.FileFormatToSaveFormat(formatInfo.FileFormatType);

                // Convert the SaveFormat to a canonical file extension (e.g., ".xlsx")
                string detectedExtension = FileFormatUtil.SaveFormatToExtension(detectedSaveFormat);

                // Get the extension present in the extracted file name
                string actualExtension = Path.GetExtension(extractedFilePath);

                // Compare extensions (case‑insensitive)
                bool isMatch = string.Equals(detectedExtension, actualExtension, StringComparison.OrdinalIgnoreCase);

                Console.WriteLine($"Original file: {Path.GetFileName(originalFilePath)}");
                Console.WriteLine($"Extracted file: {Path.GetFileName(extractedFilePath)}");
                Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");
                Console.WriteLine($"Detected extension: {detectedExtension}");
                Console.WriteLine($"Actual extension: {actualExtension}");
                Console.WriteLine($"Extension match: {isMatch}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during verification: {ex.Message}");
            }
        }

        // Example usage
        public static void Run()
        {
            // Path to the original file (could be any format)
            string originalPath = "original.xlsx";

            // Path to the extracted file that needs verification
            string extractedPath = "extracted_file.dat"; // Example: file extracted without proper extension

            // Optional: verify original file existence for completeness
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Original file not found (reference only): {originalPath}");
            }

            VerifyFileExtension(originalPath, extractedPath);
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                FileExtensionIntegrityVerifier.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
