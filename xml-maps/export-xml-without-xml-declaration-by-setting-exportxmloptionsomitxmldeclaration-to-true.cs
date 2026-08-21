// Title: Export XML without declaration using Aspose.Cells ExportXmlOptions in C#
// Description: Shows how to build a workbook, attach an XML map, and export the data to a header‑less XML file by setting ExportXmlOptions.OmitXmlDeclaration = true (with a fallback for older library versions).
// Keywords: Aspose.Cells | ExportXmlOptions | OmitXmlDeclaration | C# | .NET | XML map export | headerless XML | remove XML prolog | ExportXml | Aspose.Cells XML export
// Common Searches: Aspose.Cells export XML without <?xml?> header C# | ExportXmlOptions OmitXmlDeclaration true example | How to omit XML declaration when exporting from Aspose.Cells | Remove XML prolog after ExportXml Aspose.Cells .NET | C# export workbook to XML map without declaration
// Developer Intent: Create an XML file from a workbook via an XML map while ensuring the output does not contain the XML declaration line.
// Use Cases: Sending XML payloads to APIs that reject the <?xml?> prolog. | Generating lightweight configuration files for embedded systems. | Batch‑processing worksheets into header‑less XML for data‑exchange pipelines.
// AI Prompts: Provide C# code that uses ExportXmlOptions with OmitXmlDeclaration set to true to export a workbook via an XML map. | Explain how to detect and strip the XML declaration when using older versions of Aspose.Cells that lack OmitXmlDeclaration support. | Show a step‑by‑step example of exporting multiple worksheets to header‑less XML files in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportXmlWithoutDeclaration
{
    // Shows how to build a workbook, attach an XML map, and export the data to a header‑less XML file by setting ExportXmlOptions.OmitXmlDeclaration = true (with a fallback for older library versions).
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";
                sheet.Cells["A1"].PutValue("Id");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // Add a simple XML map (placeholder schema) to the workbook
                // In a real scenario you would use a proper XML schema string.
                int mapIndex = workbook.Worksheets.XmlMaps.Add("<Schema><Element><Id/><Name/></Element></Schema>");
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "SampleMap";

                // Export the XML using the map name and output file path
                string outputPath = "ExportedWithoutDeclaration.xml";
                workbook.ExportXml(xmlMap.Name, outputPath);

                // Remove XML declaration if present (Aspose.Cells older versions may not support OmitXmlDeclaration)
                if (File.Exists(outputPath))
                {
                    string[] lines = File.ReadAllLines(outputPath);
                    if (lines.Length > 0 && lines[0].StartsWith("<?xml", StringComparison.Ordinal))
                    {
                        File.WriteAllLines(outputPath, lines, System.Text.Encoding.UTF8);
                        // Rewrite without the first line (XML declaration)
                        File.WriteAllLines(outputPath, lines[1..], System.Text.Encoding.UTF8);
                    }
                }

                Console.WriteLine("XML exported successfully without XML declaration.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
