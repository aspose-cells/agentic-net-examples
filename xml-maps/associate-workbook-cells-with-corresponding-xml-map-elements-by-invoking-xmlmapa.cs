using System;
using Aspose.Cells;

namespace AsposeCellsXmlMappingDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Add an XML map to the workbook (using a schema file or XML string)
            // Here we assume a schema file named "schema.xsd" is available.
            int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");

            // Retrieve the added XmlMap and give it a friendly name
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "MyXmlMap";

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Associate specific cells with elements in the XML map
            // Example: link cell A1 (row 0, column 0) to the XML element "/Root/Item"
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");

            // Additional cell-to-XML mappings can be added as needed
            // sheet.Cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Root/AnotherItem");

            // Save the workbook with the XML mapping applied
            workbook.Save("output.xlsx");
        }
    }
}