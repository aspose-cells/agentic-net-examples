// Title: Extract and Save Embedded OLE Objects from an Excel Workbook with Aspose.Cells for .NET, Including Decompression and Original File Extension Preservation
// AI Prompts: Write C# code using Aspose.Cells that iterates through every worksheet, extracts each OleObject's raw byte array, checks the IsCompressed flag, decompresses the data when needed, determines the original file extension from FileExtension or Name, and writes the result to a user‑specified folder with a unique filename. | Create a reusable C# method that accepts a workbook path and an output directory, loads the workbook with Aspose.Cells, extracts all embedded OLE objects, handles possible Deflate compression, resolves the correct file extension, and saves each object as a separate file.
// Common Searches: c# aspnet extract ole objects from xlsx using aspose.cells | how to decompress embedded ole data when extracting from Excel with Aspose | save ole object to original file type aspose.cells .net | iterate worksheets to get oleobject data aspose.cells example | extract embedded pdf from excel workbook using Aspose.Cells C#
// Tags: oleobject raw byte extraction Aspose.Cells | deflate decompression of oleobject data | original file extension resolution OleObject | write oleobject bytes to filesystem | worksheet iteration for oleobjects Aspose

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet, extracts every embedded OleObject, decompresses the data if the IsCompressed flag is true, determines the original file extension via the OleObject's FileExtension or Name property, and saves each object as a uniquely named file in a specified output folder.
class OleExtractor
{
    static void Main()
    {
        // Path to the workbook that contains OLE objects
        string workbookPath = @"C:\Path\To\Input.xlsx";

        // Folder where extracted OLE files will be saved
        string outputFolder = @"C:\Path\To\OleOutput";

        try
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Verify the workbook file exists before loading
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook not found: {workbookPath}");

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            int oleIndex = 0; // Counter for unique file names

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all OLE objects on the current sheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        // Use dynamic to access members that may vary between Aspose.Cells versions
                        dynamic dynOle = ole;

                        // Retrieve the raw data of the OLE object
                        byte[] oleData = dynOle.OleObjectData as byte[];

                        // If the OLE data is compressed, decompress it
                        bool isCompressed = dynOle.IsCompressed;
                        if (isCompressed && oleData != null)
                        {
                            using (MemoryStream compressedStream = new MemoryStream(oleData))
                            using (DeflateStream deflate = new DeflateStream(compressedStream, CompressionMode.Decompress))
                            using (MemoryStream decompressedStream = new MemoryStream())
                            {
                                deflate.CopyTo(decompressedStream);
                                oleData = decompressedStream.ToArray();
                            }
                        }

                        // Determine the original file extension
                        string extension = ".bin"; // Default fallback
                        string fileExt = dynOle.FileExtension as string;
                        if (!string.IsNullOrEmpty(fileExt))
                        {
                            extension = fileExt.StartsWith(".") ? fileExt : "." + fileExt;
                        }
                        else
                        {
                            string name = dynOle.Name as string;
                            if (!string.IsNullOrEmpty(name))
                            {
                                int dotPos = name.LastIndexOf('.');
                                if (dotPos > -1 && dotPos < name.Length - 1)
                                    extension = name.Substring(dotPos);
                            }
                        }

                        // Build a unique file name for the extracted object
                        string fileName = $"OleObject_{oleIndex}{extension}";
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Save the extracted (and possibly decompressed) data to disk
                        if (oleData != null)
                            File.WriteAllBytes(outputPath, oleData);
                        else
                            Console.WriteLine($"Warning: OLE object {oleIndex} has no data.");

                        oleIndex++;
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Failed to extract an OLE object: {innerEx.Message}");
                    }
                }
            }

            Console.WriteLine($"Extraction complete. {oleIndex} OLE object(s) saved to \"{outputFolder}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
