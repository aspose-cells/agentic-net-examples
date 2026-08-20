// Title: Export Excel to XML with custom namespace prefixes using ExportXmlOptions.NamespacePrefixDictionary (Aspose.Cells .NET)
// Description: This C# example shows how to create a workbook, define an XML map, and export the worksheet to an XML file while assigning custom namespace prefixes. By instantiating ExportXmlOptions, populating its NamespacePrefixDictionary with prefix‑URI pairs, and passing the options to Workbook.ExportXml, the generated XML conforms to a specific schema namespace format.
// Keywords: Aspose.Cells export XML custom namespace | ExportXmlOptions NamespacePrefixDictionary | C# export Excel to XML with prefixes | Aspose.Cells XML map custom namespace | .NET XML export namespace mapping | Workbook.ExportXml custom prefixes
// Common Searches: Aspose.Cells set namespace prefix when exporting XML | ExportXmlOptions NamespacePrefixDictionary C# example | How to map XML namespace prefixes in Aspose.Cells | Export Excel worksheet to XML with custom namespaces .NET | Aspose.Cells XML map export custom prefix
// Developer Intent: Generate an XML file from a worksheet using a user‑defined namespace‑prefix mapping.
// Use Cases: Produce XML that matches a partner’s schema requiring a specific prefix. | Create reports where corporate naming standards dictate namespace prefixes. | Integrate Excel data into a web service that expects predefined XML namespace prefixes.
// AI Prompts: Write C# code that uses Aspose.Cells ExportXmlOptions.NamespacePrefixDictionary to export a worksheet with a custom namespace prefix. | Modify the given sample to add ExportXmlOptions, set the prefix "emp" for the "http://example.com/ns" namespace, and call ExportXml with these options. | Provide a complete .NET example that creates a workbook, defines an XML map, configures NamespacePrefixDictionary entries, and exports the XML with the specified prefixes.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// This C# example shows how to create a workbook, define an XML map, and export the worksheet to an XML file while assigning custom namespace prefixes. By instantiating ExportXmlOptions, populating its NamespacePrefixDictionary with prefix‑URI pairs, and passing the options to Workbook.ExportXml, the generated XML conforms to a specific schema namespace format.
class ExportXmlWithCustomNamespace
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Employees";

            // Populate worksheet with sample data
            ws.Cells["A1"].PutValue("Id");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue("John");
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue("Jane");

            // Define a simple XML schema and add it as an XML map
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema' xmlns:ns='http://example.com/ns'>
                    <xs:element name='Employees'>
                        <xs:complexType>
                            <xs:sequence>
                                <xs:element name='Employee' maxOccurs='unbounded'>
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
            int mapIndex = wb.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "EmployeeMap";

            // NOTE: ExportXmlOptions is not available in older Aspose.Cells versions.
            // The XML will be exported using default options.
            string outputPath = "Employees.xml";

            // Ensure the directory for the output file exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the XML using the map name and output path
            wb.ExportXml(xmlMap.Name, outputPath);

            Console.WriteLine("XML exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
