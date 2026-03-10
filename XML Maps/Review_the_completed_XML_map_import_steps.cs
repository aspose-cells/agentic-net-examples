using System;
using Aspose.Cells;

class XmlMapImportDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add an XML map from a schema file (ensure schema.xsd exists in the working directory)
        int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "SampleMap";

        // Import XML data into the first worksheet starting at cell A1 (row 0, column 0)
        workbook.ImportXml("data.xml", "Sheet1", 0, 0);

        // Link a specific cell to an element in the XML map (example: cell A2)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Root/Item");

        // Export the XML data using the defined map name
        workbook.ExportXml(xmlMap.Name, "exported.xml");

        // Save the workbook as an XLSX file
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}