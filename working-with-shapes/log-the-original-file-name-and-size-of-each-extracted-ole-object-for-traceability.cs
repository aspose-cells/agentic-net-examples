// Title: Log original file name and byte size of each extracted OLE object using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to display original file name of embedded OLE objects | retrieve size of OLE object data when extracting from Excel with Aspose.Cells | log metadata of extracted OLE objects in a .NET console application | Aspose.Cells enumerate OleObjectCollection and get ObjectSourceFullName | trace embedded OLE objects workbook processing Aspose.Cells C# example
// Tags: log OLE object source file name Aspose.Cells | retrieve embedded OLE object byte size C# | enumerate OleObjectCollection metadata .NET | extract OLE object information from Excel workbook | Aspose.Cells traceability of embedded objects

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject and OleObjectCollection

namespace OleObjectExtractor
{
    // The sample loads an Excel workbook with Aspose.Cells, iterates through each worksheet, and for every OLE object obtains the original source file name via ObjectSourceFullName and the embedded data size using ObjectData length. It writes these details to the console, providing clear traceability for each extracted OLE object; optional code for saving the binary data is included but commented out.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            try
            {
                // Validate input arguments
                if (args == null || args.Length == 0)
                {
                    Console.WriteLine("Usage: OleObjectExtractor <excel-file-path>");
                    return;
                }

                string excelPath = args[0];

                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: File not found – \"{excelPath}\"");
                    return;
                }

                // Load the workbook safely
                Workbook workbook;
                try
                {
                    workbook = new Workbook(excelPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                Console.WriteLine($"Workbook \"{Path.GetFileName(excelPath)}\" loaded successfully.");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    OleObjectCollection oleObjects = sheet.OleObjects;

                    if (oleObjects == null || oleObjects.Count == 0)
                    {
                        // No OLE objects on this sheet – continue to next
                        continue;
                    }

                    Console.WriteLine($"\nWorksheet \"{sheet.Name}\" contains {oleObjects.Count} OLE object(s):");

                    // Process each OLE object
                    for (int i = 0; i < oleObjects.Count; i++)
                    {
                        OleObject ole;
                        try
                        {
                            ole = oleObjects[i];
                        }
                        catch (Exception objEx)
                        {
                            Console.WriteLine($"  Unable to access OLE object #{i + 1}: {objEx.Message}");
                            continue;
                        }

                        // Determine original source file name (if available)
                        string originalFileName = !string.IsNullOrEmpty(ole.ObjectSourceFullName)
                            ? Path.GetFileName(ole.ObjectSourceFullName)
                            : $"EmbeddedObject_{i + 1}";

                        // Size of the embedded data (bytes). If ObjectData is null, size is 0.
                        long dataSize = ole.ObjectData?.LongLength ?? 0;

                        // Log the information
                        Console.WriteLine($"  OLE Object #{i + 1}");
                        Console.WriteLine($"    Original File Name : {originalFileName}");
                        Console.WriteLine($"    Data Size (bytes)  : {dataSize}");

                        // OPTIONAL: Save the OLE object to a file for further inspection.
                        // Uncomment the following block if you need to extract the binary data.
                        /*
                        if (ole.ObjectData != null && ole.ObjectData.Length > 0)
                        {
                            string outputDir = Path.Combine(Path.GetDirectoryName(excelPath) ?? ".", "ExtractedOleObjects");
                            Directory.CreateDirectory(outputDir);

                            string safeFileName = Path.GetInvalidFileNameChars()
                                .Aggregate(originalFileName, (current, c) => current.Replace(c, '_'));

                            string outputPath = Path.Combine(outputDir, safeFileName);
                            File.WriteAllBytes(outputPath, ole.ObjectData);
                            Console.WriteLine($"    Extracted to: {outputPath}");
                        }
                        */
                    }
                }

                Console.WriteLine("\nProcessing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
