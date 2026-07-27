using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

class ExportXmlMapsToZip
{
    static void Main()
    {
        // Path to the workbook that contains XML maps
        string workbookPath = "MappedData.xlsx";

        // Path for the resulting zip archive
        string zipPath = "ExportedXmlMaps.zip";

        // Load the workbook (uses the provided load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Verify that the workbook has XML maps
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps found in the workbook.");
            return;
        }

        // Create (or overwrite) the zip archive
        using (FileStream zipFile = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipFile, ZipArchiveMode.Update))
        {
            // Iterate through each XML map in the workbook
            for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
            {
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[i];

                // Export the XML map to a memory stream (uses the provided ExportXml rule)
                using (MemoryStream xmlStream = new MemoryStream())
                {
                    workbook.ExportXml(xmlMap.Name, xmlStream);
                    xmlStream.Position = 0; // Reset stream position for reading

                    // Define the entry name inside the zip (one file per map)
                    string entryName = $"{xmlMap.Name}.xml";

                    // Add the exported XML as a new entry in the zip archive
                    ZipArchiveEntry entry = archive.CreateEntry(entryName);
                    using (Stream entryStream = entry.Open())
                    {
                        xmlStream.CopyTo(entryStream);
                    }
                }
            }
        }

        Console.WriteLine($"Successfully exported {workbook.Worksheets.XmlMaps.Count} XML map(s) to '{zipPath}'.");
    }
}