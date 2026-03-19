using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook file
        Workbook workbook = new Workbook("input.xlsx");

        // Add an XML map to the workbook using an XSD (schema) file
        // The Add method returns the index of the newly added map
        int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        // Optionally give the map a friendly name
        xmlMap.Name = "MyXmlMap";

        // Import XML data into the first worksheet (starting at cell A1)
        // This binds the XML data to the previously added XML map
        workbook.ImportXml("data.xml", "Sheet1", 0, 0);

        // Save the workbook with the XML map and bound data
        workbook.Save("output.xlsx");
    }
}