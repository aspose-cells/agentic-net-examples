using System;
using Aspose.Cells;

class ImportXmlMapDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add an XML map to the workbook from an XSD schema file
        // The Add method returns the index of the newly added map
        int mapIndex = wb.Worksheets.XmlMaps.Add("schema.xsd");
        XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
        // Optionally give the map a friendly name
        xmlMap.Name = "MyXmlMap";

        // Import XML data into the first worksheet using the map
        // The data will be placed starting at cell A1 (row 0, column 0)
        wb.ImportXml("data.xml", "Sheet1", 0, 0);

        // Save the workbook with the imported data and XML mapping
        wb.Save("MappedWorkbook.xlsx");
    }
}