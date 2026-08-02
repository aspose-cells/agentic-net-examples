using System;
using Aspose.Cells;

class MapRangeToXml
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Define a simple XML schema for the map
        string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
            <xs:element name='Root'>
                <xs:complexType>
                    <xs:sequence>
                        <xs:element name='Item' maxOccurs='unbounded'>
                            <xs:complexType>
                                <xs:sequence>
                                    <xs:element name='Name' type='xs:string'/>
                                    <xs:element name='Value' type='xs:string'/>
                                </xs:sequence>
                            </xs:complexType>
                        </xs:element>
                    </xs:sequence>
                </xs:complexType>
            </xs:element>
        </xs:schema>";

        // Add the XML map to the workbook and give it a name
        int mapIndex = wb.Worksheets.XmlMaps.Add(xmlSchema);
        XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "RootMap";

        // Get the first worksheet and its cells collection
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate a range (A2:B3) with sample data
        cells["A2"].PutValue("Item1");
        cells["B2"].PutValue("100");
        cells["A3"].PutValue("Item2");
        cells["B3"].PutValue("200");

        // XML element path to which the cells will be linked
        string xmlPath = "/Root/Item";

        // Link each cell in the range to the XML element path
        // Row and column indices are zero‑based (A2 = row 1, column 0)
        cells.LinkToXmlMap(xmlMap.Name, 1, 0, xmlPath); // A2
        cells.LinkToXmlMap(xmlMap.Name, 1, 1, xmlPath); // B2
        cells.LinkToXmlMap(xmlMap.Name, 2, 0, xmlPath); // A3
        cells.LinkToXmlMap(xmlMap.Name, 2, 1, xmlPath); // B3

        // Save the workbook with the XML mapping applied
        wb.Save("MappedRange.xlsx");
    }
}