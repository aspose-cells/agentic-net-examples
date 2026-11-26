// Title: Export an Excel workbook's XML map to a separate .xml file while preserving the original schema using Aspose.Cells for .NET (C#)
// Common Searches: how to export an XML map from an Excel file to a .xml file using Aspose.Cells C# | Aspose.Cells ExportXml preserving original schema example | C# code to save the first XML map of a workbook as an external XML file | export xml map defined in an Excel workbook with Aspose.Cells .NET | write Excel XML map data to a separate XML file while keeping the schema in C#
// Tags: Aspose.Cells ExportXml C# | export xml map workbook Aspose.Cells | preserve original xml schema Aspose.Cells | save first xml map to external file | excel xml map extraction using .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ExportXmlMapDemo
{
    // The sample loads an Excel workbook, checks for defined XML maps, selects the first map, and uses Workbook.ExportXml to write the map's data to a separate .xml file, preserving the original schema structure.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook that contains an XML map.
            // Adjust this path to point to a real file on your system.
            string workbookPath = "input.xlsx";

            // Path where the exported XML will be saved.
            string exportPath = "exported.xml";

            // Verify that the workbook file exists.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The workbook file '{workbookPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Ensure that the workbook contains at least one XML map.
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps are defined in the workbook. Nothing to export.");
                    return;
                }

                // For this demo we export the first XML map found.
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
                string mapName = xmlMap.Name;

                Console.WriteLine($"Exporting XML map '{mapName}' to '{exportPath}'...");

                // Export the XML data using the map name.
                // This method preserves the original schema structure defined by the map.
                workbook.ExportXml(mapName, exportPath);

                // Verify that the file was created.
                if (File.Exists(exportPath))
                {
                    Console.WriteLine($"Export successful. XML file saved at: {Path.GetFullPath(exportPath)}");
                }
                else
                {
                    Console.WriteLine("Export completed but the output file was not found. Check permissions and path.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
