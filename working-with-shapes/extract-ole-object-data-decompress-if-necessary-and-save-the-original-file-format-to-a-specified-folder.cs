// Title: Extract and Decompress Embedded OLE Objects from Excel using Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans each worksheet for OleObject entries, reads FullObjectBin or ObjectData, detects Deflate or GZip compression, decompresses when needed, maps the OleObject.FileFormatType to a proper file extension, and saves the original files to a specified output folder.
// Keywords: Aspose.Cells OLE extraction | C# decompress OLE data | extract embedded objects Excel | FullObjectBin OleObject | GZip Deflate OLE detection | save OLE objects .NET | map FileFormatType to extension
// Common Searches: Aspose.Cells extract embedded OLE objects from .xlsx | C# decompress compressed OLE data in Excel | How to get original file from OleObject in Aspose.Cells | Map OleObject.FileFormatType to file extension | Save extracted OLE objects to folder using .NET
// Developer Intent: Retrieve every embedded OLE object from an Excel workbook, decompress it if compressed, and write the original file to a target directory.
// Use Cases: Bulk export of Word, PDF, image, and other embedded files for archival or migration. | Pre‑processing compressed OLE content before feeding it to analysis pipelines. | Generating an inventory of all OLE objects with sheet‑level naming for audit purposes.
// AI Prompts: Create a C# routine that iterates through all worksheets in a Workbook, extracts each OleObject's FullObjectBin or ObjectData, detects GZip or Deflate compression, decompresses the bytes, maps the FileFormatType to a common extension, and saves the file to a given folder with robust error handling. | Provide Aspose.Cells code that extracts embedded OLE objects from an Excel file, automatically handles compressed OLE streams, and preserves the original file format on disk. | Explain the logic for converting OleObject.FileFormatType values to appropriate file extensions when saving extracted OLE objects.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, scans each worksheet for OleObject entries, reads FullObjectBin or ObjectData, detects Deflate or GZip compression, decompresses when needed, maps the OleObject.FileFormatType to a proper file extension, and saves the original files to a specified output folder.
class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source workbook containing OLE objects
        string inputPath = "input.xlsx";

        // Folder where extracted files will be saved
        string outputFolder = "ExtractedOleObjects";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                OleObjectCollection oleObjects = sheet.OleObjects;

                // Process each OLE object on the current worksheet
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    // Prefer FullObjectBin (read‑only) if available; otherwise use ObjectData
                    byte[] rawData = ole.FullObjectBin ?? ole.ObjectData;

                    // Skip if there is no data
                    if (rawData == null || rawData.Length == 0)
                        continue;

                    // Determine a suitable file extension based on the object's format
                    string extension = GetExtensionFromFormat(ole.FileFormatType.ToString());

                    // Build a unique file name for the extracted object
                    string fileName = $"Sheet{sheet.Index}_Ole{i}{extension}";
                    string filePath = Path.Combine(outputFolder, fileName);

                    // Attempt to decompress the data if it appears to be compressed
                    byte[] finalData = TryDecompress(rawData);

                    // Save the extracted (and possibly decompressed) data to disk (lifecycle rule: save)
                    try
                    {
                        File.WriteAllBytes(filePath, finalData);
                        Console.WriteLine($"Saved OLE object to {filePath}");
                    }
                    catch (Exception writeEx)
                    {
                        Console.WriteLine($"Failed to write OLE object to {filePath}: {writeEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    // Maps known OLE format names to common file extensions
    private static string GetExtensionFromFormat(string formatName)
    {
        switch (formatName)
        {
            case "Word":
            case "WordDocument":
                return ".doc";
            case "Excel":
            case "ExcelWorksheet":
                return ".xls";
            case "PowerPoint":
            case "PowerPointPresentation":
                return ".ppt";
            case "Pdf":
                return ".pdf";
            case "Bitmap":
                return ".bmp";
            case "Jpeg":
                return ".jpg";
            case "Png":
                return ".png";
            case "Text":
                return ".txt";
            case "Rtf":
                return ".rtf";
            case "Html":
                return ".html";
            case "Csv":
                return ".csv";
            default:
                return ".bin";
        }
    }

    // Simple heuristic to detect and decompress common compression formats
    private static byte[] TryDecompress(byte[] data)
    {
        // Deflate (zlib) signature detection
        if (data.Length > 2 && data[0] == 0x78 && (data[1] == 0x01 || data[1] == 0x9C || data[1] == 0xDA))
        {
            try
            {
                using (var input = new MemoryStream(data))
                using (var deflate = new DeflateStream(input, CompressionMode.Decompress))
                using (var output = new MemoryStream())
                {
                    deflate.CopyTo(output);
                    return output.ToArray();
                }
            }
            catch
            {
                // If decompression fails, fall back to original data
            }
        }

        // GZip signature detection
        if (data.Length > 2 && data[0] == 0x1F && data[1] == 0x8B)
        {
            try
            {
                using (var input = new MemoryStream(data))
                using (var gzip = new GZipStream(input, CompressionMode.Decompress))
                using (var output = new MemoryStream())
                {
                    gzip.CopyTo(output);
                    return output.ToArray();
                }
            }
            catch
            {
                // If decompression fails, fall back to original data
            }
        }

        // No known compression detected; return original data
        return data;
    }
}
