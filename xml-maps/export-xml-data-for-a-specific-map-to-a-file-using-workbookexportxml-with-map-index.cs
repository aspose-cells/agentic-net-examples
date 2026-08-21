// Title: Export XML from a Selected XML Map by Index Using Aspose.Cells Workbook.ExportXml (C#)
// Description: Loads an Excel workbook, validates the requested XML‑map index, retrieves the corresponding XmlMap, and writes its XML data to a file with Workbook.ExportXml.
// Keywords: Aspose.Cells ExportXml by index | C# export XML map | Workbook.ExportXml example | Excel XML map extraction | Aspose.Cells XML map API
// Common Searches: Aspose.Cells export XML from specific map | Workbook.ExportXml with map index C# | How to get XML map name and export data | Validate XML map count before ExportXml | Export first XML map from Excel using Aspose
// Developer Intent: Write code that exports the XML data linked to a particular XML map, identified by its zero‑based index, to a file.
// Use Cases: Generate an .xml file from the first XML map in a workbook for downstream processing. | Safely export a chosen map after confirming the map index is within the XmlMaps collection range. | Log the map name and destination path after a successful ExportXml call for audit trails.
// AI Prompts: Create C# code that loads an Excel file with multiple XML maps and exports the map at a given index to a user‑specified path using Aspose.Cells. | Show how to add error handling for missing workbook files and out‑of‑range XML map indexes when calling Workbook.ExportXml. | Explain how to retrieve an XmlMap's Name from the Worksheets.XmlMaps collection and use it with ExportXml to write the XML content to disk.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportXmlByMapIndex
{
    // Loads an Excel workbook, validates the requested XML‑map index, retrieves the corresponding XmlMap, and writes its XML data to a file with Workbook.ExportXml.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel workbook that contains XML maps
            string workbookPath = "BookWithMaps.xlsx";

            // Verify that the workbook file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook wb = new Workbook(workbookPath);

                // Index of the XML map to export (change as needed)
                int mapIndex = 0;

                // Ensure the requested map index is valid
                if (wb.Worksheets.XmlMaps.Count > mapIndex)
                {
                    // Retrieve the XmlMap by its index
                    XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];

                    // Export the XML data linked to this map to a file
                    string outputPath = "ExportedMap.xml";
                    wb.ExportXml(xmlMap.Name, outputPath);

                    Console.WriteLine($"XML exported successfully to '{outputPath}' using map '{xmlMap.Name}'.");
                }
                else
                {
                    Console.WriteLine($"No XmlMap found at index {mapIndex}. Available maps: {wb.Worksheets.XmlMaps.Count}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
