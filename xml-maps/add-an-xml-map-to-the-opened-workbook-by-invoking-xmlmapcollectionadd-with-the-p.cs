using System;
using Aspose.Cells;

class AddXmlMapExample
{
    static void Main()
    {
        // Path to the existing workbook and the XML schema file
        string workbookPath = "input.xlsx";
        string schemaPath   = "schema.xsd";

        // Load the workbook (opened workbook)
        Workbook wb = new Workbook(workbookPath);

        // Access the XmlMapCollection from the workbook's worksheets
        XmlMapCollection xmlMaps = wb.Worksheets.XmlMaps;

        // Add the XML map using the schema file path
        int mapIndex = xmlMaps.Add(schemaPath);

        // Retrieve the newly added XmlMap (optional: set a friendly name)
        XmlMap xmlMap = xmlMaps[mapIndex];
        xmlMap.Name = "MyXmlMap";

        // Save the workbook with the added XML map
        wb.Save("output.xlsx");
    }
}