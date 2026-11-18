// Title: Verify that an extracted Excel workbook’s file extension matches the original using Aspose.Cells in C#
// Common Searches: c# verify extracted Excel file extension matches original Aspose.Cells | how to compare file extensions of two workbooks with Aspose.Cells .NET | check if extracted .xlsx has same extension as source file using Aspose.Cells | validate workbook file type before loading with Aspose.Cells C# | ensure Excel file integrity after extraction Aspose.Cells
// Tags: Aspose.Cells verify file extension | compare workbook extensions C# | load Excel workbook for integrity Aspose.Cells | worksheet count consistency check | handle extension mismatch .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads the original and extracted Excel files with Aspose.Cells, confirms that their file extensions are identical, and optionally validates that the worksheet counts match, throwing an exception if any discrepancy is detected.
class VerifyFileExtension
{
    static void Main()
    {
        try
        {
            // Paths to the original and the extracted files
            string originalFilePath = @"C:\Files\OriginalReport.xlsx";
            string extractedFilePath = @"C:\Files\ExtractedReport.xlsx";

            // Ensure both files exist before proceeding
            if (!File.Exists(originalFilePath))
                throw new FileNotFoundException("Original file not found.", originalFilePath);
            if (!File.Exists(extractedFilePath))
                throw new FileNotFoundException("Extracted file not found.", extractedFilePath);

            // Get file extensions (including the dot) in lower case for comparison
            string originalExtension = Path.GetExtension(originalFilePath).ToLowerInvariant();
            string extractedExtension = Path.GetExtension(extractedFilePath).ToLowerInvariant();

            // Verify that the extensions match
            if (originalExtension != extractedExtension)
                throw new InvalidOperationException(
                    $"Extension mismatch: original file is '{originalExtension}', extracted file is '{extractedExtension}'.");

            // Load both workbooks using Aspose.Cells to ensure they are valid Excel files
            Workbook originalWorkbook = new Workbook(originalFilePath);
            Workbook extractedWorkbook = new Workbook(extractedFilePath);

            // Optional: further integrity checks (e.g., compare sheet counts)
            if (originalWorkbook.Worksheets.Count != extractedWorkbook.Worksheets.Count)
                throw new InvalidOperationException("Worksheet count differs between original and extracted files.");

            // If execution reaches this point, the extension matches and both files are loadable
            Console.WriteLine("File extension verification passed and both files are valid Excel workbooks.");
        }
        catch (Exception ex)
        {
            // Output any errors encountered during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
