using System;
using Aspose.Cells;

namespace XmlMapImportDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual workbook path)
            Workbook workbook = new Workbook("input.xlsx"); // workbook-load rule

            // Path to the XML schema (XSD) that defines the mapping
            string schemaPath = "schema.xsd";

            // Add the XML map to the workbook using the schema
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath); // XmlMapCollection.Add method
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex]; // XmlMapCollection indexer

            // Optional: give the map a friendly name
            xmlMap.Name = "MyXmlMap";

            // Path to the XML data file that conforms to the schema
            string xmlDataPath = "data.xml";

            // Import the XML data into the first worksheet starting at cell A1
            // The sheet name must match an existing worksheet; here we use the first sheet's name
            string sheetName = workbook.Worksheets[0].Name; // worksheet-access rule
            workbook.ImportXml(xmlDataPath, sheetName, 0, 0); // Workbook.ImportXml method

            // Save the workbook with the imported XML map and data
            workbook.Save("output.xlsx", SaveFormat.Xlsx); // workbook-save rule
        }
    }
}