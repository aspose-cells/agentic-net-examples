// Title: Export all XML maps from an Excel workbook to individual .xml files and package them into a ZIP archive using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells, loops through its XmlMaps collection, exports each map to a MemoryStream, and adds each stream as a separate .xml entry in a ZipArchive. | Show how to build an in‑memory ZIP file with System.IO.Compression, save it to disk, and include error handling for a missing input workbook. | Demonstrate exporting XML maps directly to streams without creating temporary files, using Workbook.ExportXmlMap and streaming the results into a compressed archive.
// Common Searches: asp.net core export xml maps from excel workbook to zip using aspose.cells | c# iterate workbook xmlmaps and compress each map into a single archive | export multiple xml map definitions from an .xlsx to separate xml files without temporary files | using Aspose.Cells to create a zip of exported xml maps in memory
// Tags: Aspose.Cells export XmlMap to memory stream | C# create zip archive with multiple xml entries | Iterate workbook XmlMaps collection | Compress exported xml maps without temporary files | Workbook.ExportXmlMap usage .NET | System.IO.Compression ZipArchive for Aspose.Cells output

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The program loads an Excel workbook, verifies its existence, iterates over all defined XmlMaps, exports each map directly to a MemoryStream, adds each stream as a separate .xml entry in a ZipArchive, and writes the resulting ZIP file to disk, with basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputZip = "XmlMapsExport.zip";

            // Ensure the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that may contain XML maps
            dynamic workbook = new Workbook(inputPath);

            // Prepare a memory stream for the ZIP archive
            using (MemoryStream zipStream = new MemoryStream())
            {
                // Create a ZipArchive for adding XML files
                using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
                {
                    // Iterate through each XML map defined in the workbook (dynamic to avoid compile‑time binding)
                    foreach (dynamic xmlMap in workbook.XmlMaps)
                    {
                        // Export the current XML map to a temporary memory stream
                        using (MemoryStream xmlStream = new MemoryStream())
                        {
                            // Export the map data into the stream (no file on disk)
                            workbook.ExportXmlMap(xmlMap.Name, xmlStream);
                            xmlStream.Position = 0; // Reset stream position for reading

                            // Create a new entry in the ZIP for this map
                            ZipArchiveEntry entry = archive.CreateEntry($"{xmlMap.Name}.xml", CompressionLevel.Optimal);
                            using (Stream entryStream = entry.Open())
                            {
                                // Copy the XML content into the ZIP entry
                                xmlStream.CopyTo(entryStream);
                            }
                        }
                    }
                }

                // Write the ZIP archive to a physical file
                using (FileStream fileStream = new FileStream(outputZip, FileMode.Create, FileAccess.Write))
                {
                    zipStream.Position = 0;
                    zipStream.CopyTo(fileStream);
                }
            }

            Console.WriteLine($"XML maps have been exported to '{outputZip}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
