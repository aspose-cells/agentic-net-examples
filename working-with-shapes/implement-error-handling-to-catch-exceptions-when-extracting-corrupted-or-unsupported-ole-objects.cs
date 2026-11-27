// Title: How to add exception handling for corrupted or unsupported OLE objects when extracting them with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells extract embedded OLE objects with try‑catch for corrupted files | handle unsupported OLE format exception Aspose.Cells .NET | skip linked OLE objects and catch extraction errors using Aspose.Cells | log extraction failures of OLE objects in Excel workbook with Aspose.Cells | determine file extension from OleObject ProgID using Aspose.Cells C#
// Tags: exception handling for OLE extraction Aspose.Cells | extract embedded OLE objects C# Aspose.Cells | unsupported OLE format detection Aspose.Cells | corrupted OLE object extraction error handling | safe file name generation for extracted OLE data

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleExtractionDemo
{
    // The sample loads an Excel workbook with Aspose.Cells, creates an output folder, iterates each worksheet's OleObjectCollection, skips linked objects, and attempts to write each embedded ObjectData to a file. Workbook loading, directory creation, and each extraction are wrapped in try‑catch blocks that capture corruption or unsupported format exceptions, log the error type and message, and ensure only non‑empty data is saved with a safe file name and appropriate extension.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains OLE objects.
            // Replace with an actual file path on your machine.
            string workbookPath = @"C:\Temp\WorkbookWithOleObjects.xlsx";

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
                Console.WriteLine($"Workbook loaded successfully: {Path.GetFileName(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Directory where extracted OLE data will be saved.
            string outputDir = Path.Combine(Path.GetDirectoryName(workbookPath) ?? string.Empty, "ExtractedOleObjects");
            try
            {
                Directory.CreateDirectory(outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output directory: {ex.Message}");
                return;
            }

            // Iterate through each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"\nProcessing worksheet: {sheet.Name}");

                OleObjectCollection oleCollection = sheet.OleObjects;
                if (oleCollection == null || oleCollection.Count == 0)
                {
                    Console.WriteLine("  No OLE objects found on this sheet.");
                    continue;
                }

                // Iterate through each OLE object.
                for (int i = 0; i < oleCollection.Count; i++)
                {
                    OleObject ole = oleCollection[i];
                    Console.WriteLine($"  OLE Object Index: {i}");
                    Console.WriteLine($"    Display Label : {ole.Label}");
                    Console.WriteLine($"    ProgID        : {ole.ProgID}");
                    Console.WriteLine($"    IsLink        : {ole.IsLink}");
                    Console.WriteLine($"    FileFormatType: {ole.FileFormatType}");

                    // Skip linked OLE objects – they require external files.
                    if (ole.IsLink)
                    {
                        Console.WriteLine("    Skipped: Linked OLE objects are not extracted by this demo.");
                        continue;
                    }

                    // Attempt to extract the embedded data.
                    try
                    {
                        byte[] oleData = ole.ObjectData;

                        if (oleData == null || oleData.Length == 0)
                        {
                            Console.WriteLine("    Warning: OLE object contains no data.");
                            continue;
                        }

                        string extension = GetExtensionFromOle(ole);
                        string safeLabel = MakeFileSystemSafe(ole.Label);
                        string outputFilePath = Path.Combine(outputDir,
                            $"{sheet.Name}_Ole_{i}_{safeLabel}{extension}");

                        File.WriteAllBytes(outputFilePath, oleData);
                        Console.WriteLine($"    Extracted successfully to: {outputFilePath}");
                    }
                    catch (Exception ex)
                    {
                        // Catch any exception that indicates corruption or unsupported format.
                        Console.WriteLine($"    Extraction failed: {ex.GetType().Name} - {ex.Message}");
                    }
                }
            }

            Console.WriteLine("\nProcessing completed.");
        }

        private static string GetExtensionFromOle(OleObject ole)
        {
            // Common ProgID to extension mappings.
            string progId = ole.ProgID?.ToLowerInvariant() ?? string.Empty;

            if (progId.Contains("excel"))
                return ".xlsx";
            if (progId.Contains("word"))
                return ".docx";
            if (progId.Contains("powerpoint"))
                return ".pptx";
            if (progId.Contains("pdf"))
                return ".pdf";
            if (progId.Contains("image"))
                return ".png";

            // Fallback to FileFormatType when ProgID is not helpful.
            switch (ole.FileFormatType)
            {
                case FileFormatType.Xlsx:
                case FileFormatType.Xlsb:
                case FileFormatType.Excel97To2003:
                    return ".xlsx";

                case FileFormatType.Docx:
                case FileFormatType.Doc:
                    return ".docx";

                case FileFormatType.Pptx:
                case FileFormatType.Ppt:
                    return ".pptx";

                case FileFormatType.Pdf:
                    return ".pdf";

                case FileFormatType.Png:
                    return ".png";

                default:
                    return ".bin";
            }
        }

        /// <summary>
        /// Sanitizes a string to be safe for use as a file name.
        /// Removes or replaces characters that are invalid in Windows file names.
        /// </summary>
        private static string MakeFileSystemSafe(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "Unnamed";

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                input = input.Replace(c, '_');
            }

            // Trim length to avoid overly long file names.
            return input.Length > 50 ? input.Substring(0, 50) : input;
        }
    }
}
