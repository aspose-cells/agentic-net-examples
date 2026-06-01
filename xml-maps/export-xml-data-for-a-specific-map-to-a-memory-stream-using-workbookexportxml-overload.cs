using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class ExportXmlToMemoryStream
    {
        static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate some sample data
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

            // Link cells to the XML map so that ExportXml has data to write
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
                xmlStream.Position = 0; // Reset for reading

                // Display the exported XML content
                using (StreamReader reader = new StreamReader(xmlStream))
                {
                    string xmlContent = reader.ReadToEnd();
                    Console.WriteLine(xmlContent);
                }
            }
        }
    }
}