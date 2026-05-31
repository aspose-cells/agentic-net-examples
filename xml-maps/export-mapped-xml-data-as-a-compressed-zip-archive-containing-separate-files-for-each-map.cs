using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsExportXmlToZip
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains XML maps
            Workbook workbook = new Workbook("input.xlsx");

            // Prepare the output zip file
            string zipPath = "ExportedXmlMaps.zip";
            using (FileStream zipFileStream = new FileStream(zipPath, FileMode.Create))
            using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Create))
            {
                // Iterate through all XML maps in the workbook
                for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
                {
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[i];
                    string mapName = xmlMap.Name;

                    // Export the XML data of the current map to a memory stream
                    using (MemoryStream xmlStream = new MemoryStream())
                    {
                        workbook.ExportXml(mapName, xmlStream);
                        xmlStream.Position = 0; // Reset stream position for reading

                        // Create a zip entry named after the map (e.g., Map1.xml)
                        string entryName = $"{mapName}.xml";
                        ZipArchiveEntry entry = zipArchive.CreateEntry(entryName, CompressionLevel.Optimal);

                        // Write the exported XML into the zip entry
                        using (Stream entryStream = entry.Open())
                        {
                            xmlStream.CopyTo(entryStream);
                        }
                    }
                }
            }

            Console.WriteLine($"All XML maps have been exported to '{zipPath}'.");
        }
    }
}