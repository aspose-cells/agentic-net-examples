using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapImport
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a new workbook
            Workbook workbook = new Workbook();

            // Step 2: Load the XML schema (XSD) as an XML map
            // The XmlMaps collection is accessed via the Worksheets property
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Add the schema file to the collection; the method returns the map index
            int mapIndex = xmlMaps.Add("schema.xsd");

            // Retrieve the created XmlMap (optional: set a friendly name)
            XmlMap xmlMap = xmlMaps[mapIndex];
            xmlMap.Name = "MyXmlMap";

            // Step 3: Import XML data into the workbook using the map
            // The ImportXml method places the data starting at cell A1 of the first worksheet
            workbook.ImportXml("data.xml", "Sheet1", 0, 0);

            // Step 4: Save the workbook with the XML map applied
            workbook.Save("MappedWorkbook.xlsx");
        }
    }
}