using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Needed for OleObject and OleObjectCollection

namespace OleObjectExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input Excel file containing OLE objects
                string inputFile = @"C:\Input\WorkbookWithOleObjects.xlsx";

                // Folder where extracted files will be saved
                string outputFolder = @"C:\Output\ExtractedOleObjects";

                ExtractOleObjects(inputFile, outputFolder);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts all embedded OLE objects from the specified workbook and saves them to the given folder.
        /// </summary>
        /// <param name="workbookPath">Path to the source Excel workbook.</param>
        /// <param name="outputDir">Directory where extracted files will be written.</param>
        static void ExtractOleObjects(string workbookPath, string outputDir)
        {
            try
            {
                // Ensure the output directory exists
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Verify the workbook file exists before loading
                if (!File.Exists(workbookPath))
                {
                    Console.Error.WriteLine($"Workbook not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Iterate through each worksheet
                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    OleObjectCollection oleObjects = sheet.OleObjects;

                    // Process each OLE object in the worksheet
                    for (int oleIndex = 0; oleIndex < oleObjects.Count; oleIndex++)
                    {
                        OleObject ole = oleObjects[oleIndex];

                        // Retrieve the embedded data (byte array)
                        byte[] data = ole.ObjectData;
                        if (data == null || data.Length == 0)
                            continue; // No data to extract

                        // Determine a file name for the extracted object
                        string fileName;

                        // Prefer the original source file name if available
                        if (!string.IsNullOrEmpty(ole.ObjectSourceFullName))
                        {
                            fileName = Path.GetFileName(ole.ObjectSourceFullName);
                        }
                        else if (!string.IsNullOrEmpty(ole.SourceFullName)) // obsolete property, kept for compatibility
                        {
                            fileName = Path.GetFileName(ole.SourceFullName);
                        }
                        else
                        {
                            // Fallback to a generated name using worksheet and object indices
                            fileName = $"Sheet{wsIndex + 1}_OleObject{oleIndex + 1}.bin";
                        }

                        // Combine with output directory
                        string outputPath = Path.Combine(outputDir, fileName);

                        // Decompress if needed; otherwise use raw data
                        byte[] finalData = TryDecompress(data);

                        // Write the extracted file
                        File.WriteAllBytes(outputPath, finalData);
                        Console.WriteLine($"Extracted OLE object to: {outputPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error extracting OLE objects: {ex.Message}");
            }
        }

        /// <summary>
        /// Attempts to decompress a byte array assuming it is GZip-compressed.
        /// Returns the original array if decompression fails.
        /// </summary>
        static byte[] TryDecompress(byte[] data)
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
                // Not a GZip stream or decompression failed; return original data
                return data;
            }
        }
    }
}