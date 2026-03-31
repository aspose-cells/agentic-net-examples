using System;
using System.IO;
using Aspose.Cells;

class AddXmlMapDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Access the first worksheet and its cells
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate sample data that matches the XML schema
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("Item1");
        cells["B2"].PutValue(100);
        cells["A3"].PutValue("Item2");
        cells["B3"].PutValue(200);

        // Define an XML schema (XSD) as a string
        string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
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

        // Write the schema to a temporary file
        string tempXsdPath = Path.GetTempFileName();
        File.WriteAllText(tempXsdPath, xmlSchema);

        // Add the XML map to the workbook
        int mapIndex = wb.Worksheets.XmlMaps.Add(tempXsdPath);
        XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "DataMap";

        // Link worksheet cells to the XML map paths
        // Link the "Name" column (A2:A3) to the XML element /Data/Item/Name
        cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Data/Item/Name");
        // Link the "Value" column (B2:B3) to the XML element /Data/Item/Value
        cells.LinkToXmlMap(xmlMap.Name, 1, 1, "/Data/Item/Value");

        // Save the workbook with the XML mapping applied
        wb.Save("XmlMappedWorkbook.xlsx");

        // Clean up temporary file
        File.Delete(tempXsdPath);
    }
}