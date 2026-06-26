using System;
using Aspose.Cells;

class ImportXmlMapDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add an XML map from an XSD file to the workbook
        // The Add method returns the index of the newly added map
        int mapIndex = wb.Worksheets.XmlMaps.Add("schema.xsd");
        XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
        // Optionally give the map a friendly name
        xmlMap.Name = "MyXmlMap";

        // Import XML data into the first worksheet starting at cell A1
        // The XML data will be linked to the previously added XML map
        wb.ImportXml("data.xml", "Sheet1", 0, 0);

        // Save the workbook with the imported XML map and data
        wb.Save("MappedWorkbook.xlsx");
    }
}