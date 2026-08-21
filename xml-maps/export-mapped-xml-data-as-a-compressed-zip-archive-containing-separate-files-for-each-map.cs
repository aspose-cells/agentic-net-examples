// Title: Export All XML Maps from an Excel Workbook to a ZIP of Individual XML Files (C# Aspose.Cells)
// Description: Loads a workbook, iterates through its XmlMaps, uses Workbook.ExportXml to write each map to a MemoryStream, creates a ZipArchive entry named after the map, copies the XML data, and saves the in‑memory ZIP as ExportedXmlMaps.zip.
// Keywords: Aspose.Cells XML map export | C# ZipArchive Aspose.Cells | ExportXml multiple maps | Excel XML maps to zip | in‑memory zip C# | Workbook.ExportXml example | Aspose.Cells .NET zip archive
// Common Searches: how to export all xml maps from excel using aspose.cells | c# create zip file with each xml map from workbook | aspose.cells exportxml to ziparchive | save xml maps as separate files in a zip | aspnet stream zip of xml maps without disk
// Developer Intent: Create a single ZIP file that contains one XML file for every XML map defined in an Excel workbook.
// Use Cases: Distribute each XML map as an individual file to downstream systems. | Back up all XML map definitions from workbooks for version control. | Provide users a downloadable package of map data for reporting or data exchange.
// AI Prompts: Generate C# code that loads an Excel workbook, iterates over its XmlMaps, exports each map to a MemoryStream, and adds the XML to a ZipArchive using Aspose.Cells. | Explain how to append a timestamp to each zip entry while keeping the original map name. | Show how to stream the generated ZIP archive directly to an ASP.NET Core response without writing to disk.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsXmlExportZip
{
    // Loads a workbook, iterates through its XmlMaps, uses Workbook.ExportXml to write each map to a MemoryStream, creates a ZipArchive entry named after the map, copies the XML data, and saves the in‑memory ZIP as ExportedXmlMaps.zip.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that contains XML maps
            string workbookPath = "SourceWorkbook.xlsx";

            // Load the workbook (uses Aspose.Cells Workbook constructor)
            Workbook workbook = new Workbook(workbookPath);

            // Prepare an in‑memory stream that will hold the resulting ZIP archive
            using (MemoryStream zipStream = new MemoryStream())
            {
                // Create a ZIP archive in the memory stream
                using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
                {
                    // Iterate through all XML maps defined in the workbook
                    foreach (XmlMap xmlMap in workbook.Worksheets.XmlMaps)
                    {
                        // Export the current XML map to a temporary memory stream
                        using (MemoryStream xmlData = new MemoryStream())
                        {
                            // ExportXml(string mapName, Stream stream) – rule‑based method
                            workbook.ExportXml(xmlMap.Name, xmlData);
                            xmlData.Position = 0; // Reset position for reading

                            // Create a new entry in the ZIP archive for this map
                            ZipArchiveEntry entry = archive.CreateEntry($"{xmlMap.Name}.xml");

                            // Write the exported XML data into the ZIP entry
                            using (Stream entryStream = entry.Open())
                            {
                                xmlData.CopyTo(entryStream);
                            }
                        }
                    }
                }

                // Save the ZIP archive to a physical file
                File.WriteAllBytes("ExportedXmlMaps.zip", zipStream.ToArray());
            }

            Console.WriteLine("All XML maps have been exported to ExportedXmlMaps.zip");
        }
    }
}
