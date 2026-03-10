using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapPrerequisiteDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Sample XSD and XML content
            string xsdContent = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' type='xs:string' minOccurs='0' maxOccurs='unbounded'/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            string xmlContent = @"<?xml version='1.0' encoding='utf-8'?>
<Root>
  <Item>Value1</Item>
  <Item>Value2</Item>
</Root>";

            // Write XSD and XML to temporary files
            string tempXsdPath = Path.Combine(Path.GetTempPath(), "sample.xsd");
            string tempXmlPath = Path.Combine(Path.GetTempPath(), "sample.xml");
            File.WriteAllText(tempXsdPath, xsdContent);
            File.WriteAllText(tempXmlPath, xmlContent);

            // Add an XML map to the workbook using the schema file
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "SampleMap";

            // Import the XML data into the first worksheet starting at cell A1
            workbook.ImportXml(tempXmlPath, "Sheet1", 0, 0);

            // Save the workbook to an Excel file
            workbook.Save("ImportedXml.xlsx", SaveFormat.Xlsx);
        }
    }
}