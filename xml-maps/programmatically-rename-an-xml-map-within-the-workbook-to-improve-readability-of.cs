using System;
using System.IO;
using Aspose.Cells;

class RenameXmlMapDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define a simple XML schema
        string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                <xs:element name='Employee'>
                                    <xs:complexType>
                                        <xs:sequence>
                                            <xs:element name='ID' type='xs:int'/>
                                            <xs:element name='Name' type='xs:string'/>
                                        </xs:sequence>
                                    </xs:complexType>
                                </xs:element>
                             </xs:schema>";

        // Save schema to a temporary file
        string schemaPath = Path.Combine(Path.GetTempPath(), "EmployeeSchema.xsd");
        File.WriteAllText(schemaPath, xmlSchema);

        // Add the XML map to the workbook
        int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

        // Output the default name (usually the root element name)
        Console.WriteLine("Original XML Map Name: " + xmlMap.Name);

        // Rename the XML map to a more readable identifier
        xmlMap.Name = "EmployeeDataMap";

        // Verify that the name has been changed
        Console.WriteLine("Renamed XML Map Name: " + xmlMap.Name);

        // Export XML using the new map name (demonstration purpose)
        workbook.ExportXml(xmlMap.Name, "EmployeeData.xml");

        // Save the workbook with the renamed XML map
        workbook.Save("RenamedXmlMapWorkbook.xlsx");
    }
}