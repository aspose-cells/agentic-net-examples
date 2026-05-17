using System;
using System.IO;
using Aspose.Cells;

class LinkCellToXmlMapDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Define a simple XML schema (XSD) as a string
        string xmlSchema = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
    <xs:element name='Root'>
        <xs:complexType>
            <xs:sequence>
                <xs:element name='Value' type='xs:string'/>
            </xs:sequence>
        </xs:complexType>
    </xs:element>
</xs:schema>";

        // Write the schema to a temporary file
        string tempXsdPath = Path.GetTempFileName();
        File.WriteAllText(tempXsdPath, xmlSchema);

        // Add the XML map to the workbook and retrieve it
        int mapIndex = wb.Worksheets.XmlMaps.Add(tempXsdPath);
        XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "MyMap";

        // Link cell A1 (row 0, column 0) to the XML element "/Root/Value"
        wb.Worksheets[0].Cells.LinkToXmlMap("MyMap", 0, 0, "/Root/Value");

        // Optional: set an initial value in the linked cell
        wb.Worksheets[0].Cells["A1"].PutValue("Sample");

        // Save the workbook
        wb.Save("LinkedCell.xlsx");

        // Clean up temporary file
        File.Delete(tempXsdPath);
    }
}