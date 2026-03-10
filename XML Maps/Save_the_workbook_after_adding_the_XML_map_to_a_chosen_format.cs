using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Fill some sample data
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Laptop");
        worksheet.Cells["B2"].PutValue(999.99);

        // Define XML schema
        string xmlSchema = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""ProductData"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""Product"" type=""xs:string""/>
        <xs:element name=""Price"" type=""xs:double""/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // Save schema to a temporary file
        string schemaPath = Path.Combine(Path.GetTempPath(), "ProductData.xsd");
        File.WriteAllText(schemaPath, xmlSchema);

        // Add an XML map to the workbook using the schema file
        int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "ProductDataMap";

        // Configure XmlSaveOptions to export the XML map
        XmlSaveOptions saveOptions = new XmlSaveOptions
        {
            XmlMapName = xmlMap.Name
        };

        // Save the workbook as an XML file with the map included
        workbook.Save("ProductData.xml", saveOptions);
    }
}