using System;
using Aspose.Cells;

class ImportXmlMapDemo
{
    static void Main()
    {
        // Load an existing workbook (any supported format, e.g., .xlsx)
        Workbook workbook = new Workbook("template.xlsx"); // load rule

        // Add an XML map to the workbook using an XSD schema file
        int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd"); // add rule

        // Optionally set a friendly name for the map
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "MyXmlMap";

        // Import XML data into the first worksheet starting at cell A1
        workbook.ImportXml("data.xml", "Sheet1", 0, 0); // import rule

        // Save the workbook with the imported XML map and data
        workbook.Save("output.xlsx"); // save rule
    }
}