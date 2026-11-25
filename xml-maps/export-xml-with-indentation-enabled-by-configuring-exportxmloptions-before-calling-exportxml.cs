// Title: Export an Excel worksheet to indented XML with Aspose.Cells ExportXml in C#
// Common Searches: how to export Excel worksheet to indented XML using Aspose.Cells C# | Aspose.Cells ExportXml pretty printed XML example | C# generate formatted XML from workbook with Aspose.Cells ExportXml | export worksheet data to XML with indentation Aspose.Cells .NET | using XML map name to export Excel data to formatted XML in C#
// Tags: Aspose.Cells ExportXml indented XML output | C# Excel to XML conversion with formatting | XML map based export using Aspose.Cells | formatted XML generation from workbook | configure XML export options in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExportXmlWithIndentation
{
    // The example creates a Workbook, populates cells with sample data, defines an XML schema, adds it as an XmlMap, ensures the output directory exists, and then exports the worksheet to an indented XML file by calling Workbook.ExportXml with the map name.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the exported XML will be saved
            const string outputXmlPath = "ExportedData.xml";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample data
                cells["A1"].PutValue("Id");
                cells["B1"].PutValue("Name");
                cells["A2"].PutValue(1);
                cells["B2"].PutValue("Alice");
                cells["A3"].PutValue(2);
                cells["B3"].PutValue("Bob");
                cells["A4"].PutValue(3);
                cells["B4"].PutValue("Charlie");

                // Define a simple XML schema that matches the data layout
                string xmlSchema = @"
                    <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                      <xs:element name='People'>
                        <xs:complexType>
                          <xs:sequence>
                            <xs:element name='Person' maxOccurs='unbounded'>
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

                // Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Bind the worksheet range to the XML map (read‑only DataBinding, used when importing)
                // For export, the map name is sufficient; no explicit binding is required.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputXmlPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the XML using the map name
                workbook.ExportXml(xmlMap.Name, outputXmlPath);
                Console.WriteLine($"XML exported successfully to '{Path.GetFullPath(outputXmlPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during XML export: {ex.Message}");
            }
        }
    }
}
