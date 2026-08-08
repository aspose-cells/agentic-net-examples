// Title: Export a Workbook to XML with a Custom Root Element using Aspose.Cells C#
// Description: Shows how to build an Aspose.Cells workbook, create an XSD schema that defines a custom root node, add the schema as an XmlMap, and call ExportXml to produce an XML file with that root element in C#.
// Keywords: Aspose.Cells ExportXml | C# export workbook to XML | custom root element XML Aspose | XmlMap XSD schema | temporary XSD file Aspose.Cells | Excel to XML with custom root | Aspose.Cells XML map example | Export workbook as XML C# | Define XML schema for Excel export | Aspose.Cells XML export guide
// Common Searches: Aspose.Cells export XML custom root | How to use XmlMap with XSD in Aspose.Cells | C# export Excel to XML with specific root element | Add temporary XSD to workbook Aspose.Cells | ExportXml with custom root node | Aspose.Cells XML map example C# | Create XML from workbook using XSD schema
// Developer Intent: Produce an XML document from a workbook where the root element is defined by a user‑supplied XSD map.
// Use Cases: Integrate Excel data into systems that require a predefined XML structure with a specific root tag. | Generate XML reports that conform to an external XSD without modifying the original workbook layout. | Automate data exchange between .NET applications and services that consume custom‑root XML payloads.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to XML using an XSD that sets a custom root element. | Explain the steps to add an XmlMap from a temporary XSD file and call ExportXml for a custom root node. | Provide troubleshooting tips when ExportXml fails because the XSD file is missing or incorrectly formatted.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to build an Aspose.Cells workbook, create an XSD schema that defines a custom root node, add the schema as an XmlMap, and call ExportXml to produce an XML file with that root element in C#.
class ExportWorkbookWithCustomRoot
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define an XML schema whose root element is "CustomRoot"
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                    <xs:element name='CustomRoot'>
                        <xs:complexType>
                            <xs:sequence>
                                <xs:element name='Record' maxOccurs='unbounded'>
                                    <xs:complexType>
                                        <xs:sequence>
                                            <xs:element name='Id' type='xs:integer'/>
                                            <xs:element name='Name' type='xs:string'/>
                                        </xs:sequence>
                                    </xs:complexType>
                                </xs:element>
                            </xs:sequence>
                        </xs:complexType>
                    </xs:element>
                </xs:schema>";

            // Write the schema to a temporary file to satisfy XmlMaps.Add(string filePath)
            string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
            File.WriteAllText(tempSchemaPath, xmlSchema);

            // Ensure the temporary schema file exists before adding
            if (!File.Exists(tempSchemaPath))
                throw new FileNotFoundException("Temporary XML schema file was not created.", tempSchemaPath);

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "CustomMap";

            // Export the workbook data to XML using the custom map
            string outputPath = "CustomRootOutput.xml";
            workbook.ExportXml(xmlMap.Name, outputPath);

            Console.WriteLine("XML exported successfully with custom root element.");

            // Clean up temporary schema file
            if (File.Exists(tempSchemaPath))
                File.Delete(tempSchemaPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
