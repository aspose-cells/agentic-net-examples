// Title: Export XML Map to MemoryStream using Aspose.Cells Workbook.ExportXml (C#)
// Description: Demonstrates how to create a workbook, define an XML schema, map worksheet cells, and export the linked XML directly into a MemoryStream with the Workbook.ExportXml overload. The stream is reset for immediate reading, enabling in‑memory XML handling without writing a file.
// Keywords: Aspose.Cells ExportXml | XML map to MemoryStream | C# Aspose.Cells example | Workbook.ExportXml overload | in‑memory XML export .NET | Aspose.Cells XML mapping | export worksheet data as XML stream
// Common Searches: Aspose.Cells export XML map to MemoryStream | Workbook.ExportXml C# example | How to write XML map data to a stream with Aspose.Cells | Export linked worksheet cells as XML without saving file | Aspose.Cells XML map memory stream .NET
// Developer Intent: Write the XML produced by a specific XML map directly to a MemoryStream.
// Use Cases: Generate an XML payload for a web API without creating a temporary file. | Pass in‑memory XML to another service or component in a data pipeline. | Validate or transform exported XML before deciding to persist it.
// AI Prompts: Show a C# code snippet that uses Aspose.Cells Workbook.ExportXml to export a named XML map into a MemoryStream and returns the XML string. | Explain the steps to link worksheet cells to an XML map and then export the mapped XML directly to a stream with Aspose.Cells. | Provide guidance on handling exceptions and resetting the MemoryStream position after exporting XML using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, define an XML schema, map worksheet cells, and export the linked XML directly into a MemoryStream with the Workbook.ExportXml overload. The stream is reset for immediate reading, enabling in‑memory XML handling without writing a file.
class ExportXmlToMemoryStream
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define a simple XML schema and add it as an XML map
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence>
                                                            <xs:element name='Id' type='xs:int'/>
                                                            <xs:element name='Name' type='xs:string'/>
                                                        </xs:sequence>
                                                    </xs:complexType>
                                                </xs:element>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "SampleMap";

            // Link worksheet cells to the XML map paths
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item/Id");
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 1, "/Root/Item/Name");
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Root/Item/Id");
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 1, 1, "/Root/Item/Name");
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 2, 0, "/Root/Item/Id");
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 2, 1, "/Root/Item/Name");

            // Export the XML data linked by the map to a memory stream
            using (MemoryStream xmlStream = new MemoryStream())
            {
                workbook.ExportXml(xmlMap.Name, xmlStream);
                xmlStream.Position = 0; // Reset stream position for reading

                // Display the exported XML content
                using (StreamReader reader = new StreamReader(xmlStream))
                {
                    string xmlContent = reader.ReadToEnd();
                    Console.WriteLine("Exported XML:");
                    Console.WriteLine(xmlContent);
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        ExportXmlToMemoryStream.Run();
    }
}
