// Title: Export XML from an Aspose.Cells workbook without the XML declaration using C#
// Common Searches: how to export XML from Aspose.Cells workbook without XML declaration in C# | Aspose.Cells ExportXml remove XML prolog .NET | C# export workbook to XML without <?xml?> header using Aspose.Cells | remove XML declaration when exporting XML map with Aspose.Cells | export XML map without XML declaration using Aspose.Cells C#
// Tags: Aspose.Cells export XML without declaration | C# remove XML prolog after ExportXml | XML map export without XML header | Workbook ExportXml without XML declaration | Aspose.Cells XML export without prolog

using System;
using System.IO;
using Aspose.Cells;

namespace ExportXmlWithoutDeclaration
{
    // The example creates a Workbook, fills it with sample data, adds an XML map based on an XSD schema, and exports the data to an XML file. After export it detects the XML declaration line and rewrites the file without that line, resulting in an XML output that lacks the <?xml?> prolog.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputXmlPath = "ExportedWithoutDeclaration.xml";

            // Sample XML schema defining the structure of the XML to be exported.
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                  <xs:element name='Data'>
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element name='Item' maxOccurs='unbounded'>
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name='Name' type='xs:string'/>
                              <xs:element name='Value' type='xs:integer'/>
                            </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>
                </xs:schema>";

            try
            {
                // 1. Create a new workbook and populate it with sample data.
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Header row
                    cells["A1"].PutValue("Name");
                    cells["B1"].PutValue("Value");

                    // Data rows
                    cells["A2"].PutValue("Item1");
                    cells["B2"].PutValue(100);
                    cells["A3"].PutValue("Item2");
                    cells["B3"].PutValue(200);
                    cells["A4"].PutValue("Item3");
                    cells["B4"].PutValue(300);

                    // 2. Add the XML map to the workbook using the schema defined above.
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                    // 3. Export the XML using the map name.
                    workbook.ExportXml(xmlMap.Name, outputXmlPath);
                }

                // 4. Remove the XML declaration if it exists.
                if (File.Exists(outputXmlPath))
                {
                    string[] lines = File.ReadAllLines(outputXmlPath);
                    if (lines.Length > 0 && lines[0].StartsWith("<?xml", StringComparison.Ordinal))
                    {
                        File.WriteAllLines(outputXmlPath, lines, System.Text.Encoding.UTF8);
                        // Rewrite without the first line (the declaration)
                        File.WriteAllText(outputXmlPath, string.Join(Environment.NewLine, lines, 1, lines.Length - 1));
                    }

                    Console.WriteLine($"XML exported successfully to '{outputXmlPath}'.");
                    Console.WriteLine("First 200 characters of the file (for verification):");
                    string preview = File.ReadAllText(outputXmlPath);
                    Console.WriteLine(preview.Length > 200 ? preview.Substring(0, 200) + "..." : preview);
                }
                else
                {
                    Console.WriteLine("Export completed but the output file was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
