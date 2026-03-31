using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing Excel workbook
        string workbookPath = "input.xlsx";

        // Path to the XML schema (XSD) file
        string schemaPath = "schema.xsd";

        // Load the workbook from the file
        Workbook workbook = new Workbook(workbookPath);

        // Add an XML map to the workbook using the schema file
        int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

        // Optionally assign a friendly name to the XML map
        xmlMap.Name = "MyXmlMap";

        // Save the workbook with the associated XML map
        workbook.Save("output_with_map.xlsx");
    }
}