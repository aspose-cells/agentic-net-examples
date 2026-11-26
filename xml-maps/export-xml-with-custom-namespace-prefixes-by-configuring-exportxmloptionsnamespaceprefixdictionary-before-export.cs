// Title: Export workbook data to XML with custom namespace prefixes using ExportXmlOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells export XML with custom namespace prefixes using ExportXmlOptions | C# set NamespacePrefixDictionary before exporting XML map in Aspose.Cells | Configure ExportXmlOptions to map XML namespaces in .NET workbook export | Export Excel worksheet to XML with multiple namespaces Aspose.Cells example | How to apply custom namespace prefixes when exporting XML from a workbook in C#
// Tags: ExportXmlOptions NamespacePrefixDictionary Aspose.Cells | custom XML namespace prefixes C# | XML map export with custom namespaces .NET | Aspose.Cells XML export options configuration | namespace prefix dictionary for XML export

using System;
using System.IO;
using Aspose.Cells;

namespace ExportXmlWithCustomNamespaces
{
    // This example demonstrates how to configure ExportXmlOptions.NamespacePrefixDictionary to assign custom namespace prefixes, add an XML map based on a schema, and export worksheet data to an XML file with the specified prefixes using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "ExportedData.xml";

            try
            {
                // Create a new workbook and get the first worksheet
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Populate sample data
                    cells["A1"].PutValue("Id");
                    cells["B1"].PutValue("Name");
                    cells["A2"].PutValue(1);
                    cells["B2"].PutValue("Alice");
                    cells["A3"].PutValue(2);
                    cells["B3"].PutValue("Bob");

                    // XML schema defining the map
                    string xmlSchema = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'
           xmlns:ex='http://example.com/ns1'
           xmlns:ex2='http://example.com/ns2'
           targetNamespace='http://example.com/ns1'
           elementFormDefault='qualified'>
  <xs:element name='Root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Id' type='xs:int' />
              <xs:element name='Name' type='xs:string' />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                    // Add the XML map to the workbook
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                    // Verify the map query (optional)
                    sheet.GetAreasOfXmlMapQuery("/ex:Root/ex:Item", xmlMap);

                    // Export the XML using the map name
                    workbook.ExportXml(xmlMap.Name, outputPath);

                    Console.WriteLine($"XML exported successfully to '{Path.GetFullPath(outputPath)}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during XML export:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
