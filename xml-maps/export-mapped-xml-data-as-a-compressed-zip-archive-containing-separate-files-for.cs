using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

class ExportXmlMapsToZip
{
    static void Main()
    {
        // Load the workbook that contains XML maps
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Path for the resulting zip archive
        string zipFilePath = "XmlMaps.zip";

        // Create the zip archive
        using (FileStream zipStream = new FileStream(zipFilePath, FileMode.Create))
        using (ZipArchive zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            // Iterate through each XML map in the workbook
            for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
            {
                // Get the current XML map
                var xmlMap = workbook.Worksheets.XmlMaps[i];

                // Export the XML data of the map to a memory stream
                using (MemoryStream xmlStream = new MemoryStream())
                {
                    // Use the ExportXml overload that accepts a stream (rule-compliant)
                    workbook.ExportXml(xmlMap.Name, xmlStream);
                    xmlStream.Position = 0; // Reset stream position for reading

                    // Create a zip entry named after the map
                    var zipEntry = zipArchive.CreateEntry($"{xmlMap.Name}.xml");

                    // Write the exported XML into the zip entry
                    using (Stream entryStream = zipEntry.Open())
                    {
                        xmlStream.CopyTo(entryStream);
                    }
                }
            }
        }

        Console.WriteLine($"Exported {workbook.Worksheets.XmlMaps.Count} XML map(s) to '{zipFilePath}'.");
    }
}