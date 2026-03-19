using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (any supported Excel format)
        Workbook workbook = new Workbook("input.xlsx");

        // Add an XML map to the workbook using a schema (XSD) file path
        // The Add method returns the index of the newly added map
        int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");

        // Retrieve the XmlMap object and give it a friendly name
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "MyXmlMap";

        // Get the first worksheet where we will associate cells with XML elements
        Worksheet sheet = workbook.Worksheets[0];

        // Link cell A1 (row 0, column 0) to an XML element path in the map
        // Adjust the path to match the schema's element hierarchy
        sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");

        // Optionally link more cells as needed
        // sheet.Cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Root/AnotherItem");

        // Save the workbook with the XML mapping applied
        workbook.Save("output.xlsx");
    }
}