// Title: C# – Extract and Save Embedded OLE Objects from Excel with Aspose.Cells (auto‑decompress)
// Description: Loads an Excel workbook, walks through every worksheet, pulls each embedded OLE object, determines its original extension via FileFormatType, transparently decompresses GZip data when needed, and writes the file to a user‑defined folder with a unique name.
// Keywords: Aspose.Cells | C# | OLE object extraction | Excel embedded files | GZip decompression | FileFormatType mapping | save OLE to disk | batch OLE export | extract embedded Word PDF | auto‑decompress OLE
// Common Searches: extract OLE objects from Excel using Aspose.Cells C# | decompress embedded OLE data in .NET | map Aspose.Cells FileFormatType to file extension | save extracted OLE files to folder | batch export embedded documents from workbook
// Developer Intent: Retrieve every embedded OLE object from an Excel file, decompress it if compressed, and store it with the correct file extension.
// Use Cases: Archive all Word, PDF, and image files embedded across multiple sheets for compliance. | Automate a document‑ingestion pipeline that extracts, decompresses, and indexes OLE content in a DMS. | Generate a catalog of embedded objects, including source sheet, inferred type, and saved path.
// AI Prompts: Write C# code with Aspose.Cells that extracts OLE objects, determines their extensions, auto‑decompresses GZip data, and saves them to a target directory. | Explain a safe way to convert Aspose.Cells FileFormatType values to common file extensions without hard‑coding version‑specific enum members. | Suggest improvements for error handling, unknown format fallback, and preserving original OLE object names during extraction.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject and OleObjectCollection

namespace OleObjectExtractor
{
    // Loads an Excel workbook, walks through every worksheet, pulls each embedded OLE object, determines its original extension via FileFormatType, transparently decompresses GZip data when needed, and writes the file to a user‑defined folder with a unique name.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file that contains OLE objects
            string excelFilePath = @"C:\Input\WorkbookWithOleObjects.xlsx";

            // Folder where extracted files will be saved
            string outputFolder = @"C:\Output\ExtractedOleObjects";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Verify the input file exists before loading
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Input file not found: {excelFilePath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(excelFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through all worksheets
            for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIdx];
                OleObjectCollection oleObjects = sheet.OleObjects;

                // Process each OLE object in the current worksheet
                for (int oleIdx = 0; oleIdx < oleObjects.Count; oleIdx++)
                {
                    OleObject ole = oleObjects[oleIdx];

                    // Retrieve the embedded data. Prefer ObjectData; if null, fall back to FullObjectBin.
                    byte[] data = ole.ObjectData ?? ole.FullObjectBin;

                    if (data == null || data.Length == 0)
                    {
                        Console.WriteLine($"Worksheet {sheetIdx}, OLE {oleIdx}: No data found.");
                        continue;
                    }

                    // Determine file extension based on the FileFormatType property
                    string extension = GetExtensionFromFormat(ole.FileFormatType);

                    // Build a unique file name
                    string fileName = $"Sheet{sheetIdx}_Ole{oleIdx}{extension}";
                    string outputPath = Path.Combine(outputFolder, fileName);

                    // Attempt to decompress if the data is compressed
                    byte[] finalData = TryDecompress(data);

                    try
                    {
                        // Save the extracted file
                        File.WriteAllBytes(outputPath, finalData);
                        Console.WriteLine($"Extracted OLE object to: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to write file '{outputPath}': {ex.Message}");
                    }
                }
            }
        }

        // Maps Aspose.Cells FileFormatType to a typical file extension.
        private static string GetExtensionFromFormat(FileFormatType format)
        {
            // Use the enum name as a string to avoid version‑specific enum members.
            string name = format.ToString().ToLowerInvariant();

            switch (name)
            {
                case "doc":
                case "docx":
                    return ".docx";

                case "xls":
                case "xlsx":
                case "xlsb":
                case "xlsm":
                    return ".xlsx";

                case "ppt":
                case "pptx":
                    return ".pptx";

                case "pdf":
                    return ".pdf";

                case "txt":
                    return ".txt";

                case "rtf":
                    return ".rtf";

                case "html":
                case "mhtml":
                    return ".html";

                case "csv":
                    return ".csv";

                default:
                    return ".bin"; // Fallback for unknown types
            }
        }

        // Attempts to decompress data assuming it might be GZip-compressed.
        // Returns the original data if decompression is not applicable.
        private static byte[] TryDecompress(byte[] data)
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
                // If decompression fails, assume data was not compressed.
                return data;
            }
        }
    }
}
