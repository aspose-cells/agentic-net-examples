// Title: Export all XML maps from an Excel workbook to separate .xml files and compress them into a zip archive using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each XML map to its own .xml file and add to a zip archive | How to iterate XmlMapCollection and create a zip file with the exported XML using Aspose.Cells | Save multiple Excel XML maps as individual files in a compressed zip with Aspose.Cells .NET | Export XmlMap to MemoryStream and zip the results programmatically in C# | Sanitize XML map names for safe zip entry names when exporting with Aspose.Cells
// Tags: export xml maps Aspose.Cells | Workbook.ExportXml memory stream | zip archive creation C# | safe zip entry naming for XML maps | iterate XmlMapCollection Aspose.Cells | compress multiple xml files .NET

using System;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The sample loads an Excel workbook, loops through its XmlMapCollection, uses Workbook.ExportXml to write each map to a MemoryStream, sanitizes the map name for a valid file name, and stores each XML stream as a separate entry in a newly created zip archive.
class Program
{
    private static string SanitizeFileName(string name)
    {
        // Replace any invalid characters with an underscore
        string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
        string pattern = $"[{Regex.Escape(invalidChars)}]";
        return Regex.Replace(name, pattern, "_");
    }

    static void Main(string[] args)
    {
        // Determine input workbook and output zip paths
        string inputWorkbookPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputZipPath = args.Length > 1 ? args[1] : "mapped_xml.zip";

        // Validate input workbook existence
        if (!File.Exists(inputWorkbookPath))
        {
            Console.Error.WriteLine($"Error: Input workbook not found at '{inputWorkbookPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            using (Workbook workbook = new Workbook(inputWorkbookPath))
            {
                // Access the collection of XML maps
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                if (xmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps are defined in the workbook. No files will be exported.");
                    return;
                }

                // Create (or overwrite) the zip archive
                using (FileStream zipFileStream = new FileStream(outputZipPath, FileMode.Create, FileAccess.Write))
                using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Create))
                {
                    foreach (XmlMap map in xmlMaps)
                    {
                        // Export each map to a memory stream
                        using (MemoryStream xmlStream = new MemoryStream())
                        {
                            // ExportXml writes the XML representation of the map into the stream
                            workbook.ExportXml(map.Name, xmlStream);
                            xmlStream.Position = 0; // Reset stream position for reading

                            // Prepare a safe entry name
                            string entryFileName = $"{SanitizeFileName(map.Name)}.xml";

                            // Create a new entry in the zip archive
                            ZipArchiveEntry entry = zipArchive.CreateEntry(entryFileName, CompressionLevel.Optimal);
                            using (Stream entryStream = entry.Open())
                            {
                                xmlStream.CopyTo(entryStream);
                            }

                            Console.WriteLine($"Exported XML map '{map.Name}' to zip entry '{entryFileName}'.");
                        }
                    }
                }

                Console.WriteLine($"All XML maps have been exported and compressed into '{outputZipPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
