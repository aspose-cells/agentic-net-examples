using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapImport
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel workbook
            string workbookPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path to the XML schema (XSD) that defines the XML map
            string schemaPath = "schema.xsd";

            // Add the XML map to the workbook using the schema file
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Optionally set a friendly name for the map
            xmlMap.Name = "MyXmlMap";

            // Save the workbook with the newly added XML map
            workbook.Save("output.xlsx");
        }
    }
}