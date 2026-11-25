// Title: Export a workbook to UTF‑8 encoded XML using Aspose.Cells ExportXml in C#
// Common Searches: how to export a worksheet to UTF-8 XML with Aspose.Cells in C# | Aspose.Cells ExportXml method write UTF-8 encoded file | C# export data to XML map using Aspose.Cells with specific encoding | using FileStream with Aspose.Cells ExportXml to generate UTF-8 XML | Aspose.Cells XML map export example C# UTF-8
// Tags: Aspose.Cells ExportXml UTF-8 | C# export XML map Aspose.Cells | Workbook ExportXml to file | XML map schema Aspose.Cells | FileStream Aspose.Cells ExportXml

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace ExportXmlUtf8Demo
{
    // The example creates a workbook, defines an XML schema, adds it as an XML map, and uses Aspose.Cells ExportXml with a FileStream to generate a UTF‑8 encoded XML file.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputXmlPath = "ExportedData.xml";

            // XML schema defining the structure of the XML map
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
                // Create a new workbook and populate it with sample data
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "SampleData";

                    // Header row
                    sheet.Cells["A1"].PutValue("Name");
                    sheet.Cells["B1"].PutValue("Value");

                    // Data rows
                    sheet.Cells["A2"].PutValue("Item1");
                    sheet.Cells["B2"].PutValue(100);
                    sheet.Cells["A3"].PutValue("Item2");
                    sheet.Cells["B3"].PutValue(200);
                    sheet.Cells["A4"].PutValue("Item3");
                    sheet.Cells["B4"].PutValue(300);

                    // Add the XML map to the workbook using the schema defined above
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                    // Export the XML using a FileStream (Aspose.Cells writes UTF‑8 by default)
                    using (FileStream fs = new FileStream(outputXmlPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        workbook.ExportXml(xmlMap.Name, fs);
                    }

                    // Verify that the file is UTF‑8 encoded by reading the first line
                    using (StreamReader reader = new StreamReader(outputXmlPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                    {
                        string firstLine = reader.ReadLine();
                        Console.WriteLine("First line of exported XML:");
                        Console.WriteLine(firstLine);
                    }

                    Console.WriteLine($"XML exported successfully to '{outputXmlPath}' with UTF‑8 encoding.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
