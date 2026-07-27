using System;
using System.IO;
using System.Xml.Linq;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonExportDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and populate sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Id");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // 2. Define a simple XML schema that maps the worksheet data
                string xmlSchema = @"
                    <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                      <xs:element name='Items'>
                        <xs:complexType>
                          <xs:sequence>
                            <xs:element name='Item' maxOccurs='unbounded'>
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

                // 3. Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "ItemMap";

                // 4. Export the mapped data to XML using a memory stream
                using (MemoryStream xmlStream = new MemoryStream())
                {
                    workbook.ExportXml(xmlMap.Name, xmlStream);
                    xmlStream.Position = 0; // Reset stream position for reading

                    // 5. Load the exported XML into an XDocument
                    XDocument xDoc = XDocument.Load(xmlStream);

                    // 6. Convert the XDocument to JSON (preserving the XML hierarchy)
                    string json = JsonSerializer.Serialize(
                        xDoc,
                        new JsonSerializerOptions { WriteIndented = true });

                    // 7. Save the JSON string to a file
                    string jsonPath = "MappedData.json";
                    File.WriteAllText(jsonPath, json);

                    Console.WriteLine($"JSON file generated successfully at: {Path.GetFullPath(jsonPath)}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}