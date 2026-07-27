using System;
using Aspose.Cells;

class ImportXmlAndLinkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Import XML data into the first worksheet starting at cell A1
        wb.ImportXml("data.xml", "Sheet1", 0, 0);

        // Add an XML map using a schema file (XSD)
        int mapIndex = wb.Worksheets.XmlMaps.Add("schema.xsd");
        XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "DataMap";

        // Get the first worksheet and its cells collection
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Link cells to specific XML map elements
        // Example: link cell A2 (row 1, column 0) to the XML element /Root/Item/Name
        cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Root/Item/Name");
        // Example: link cell B2 (row 1, column 1) to the XML element /Root/Item/Price
        cells.LinkToXmlMap(xmlMap.Name, 1, 1, "/Root/Item/Price");

        // Save the workbook with the linked XML map
        wb.Save("LinkedXmlWorkbook.xlsx");
    }
}