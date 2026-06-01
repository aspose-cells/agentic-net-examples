using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel workbook (.xlsx)
            string excelPath = "input.xlsx";

            // Load the workbook from the .xlsx file
            Workbook workbook = new Workbook(excelPath);

            // Path or content of the XML schema (XSD) to be attached as a map
            // You can provide a file path, e.g., "schema.xsd", or the schema string itself.
            string xmlSchemaPath = "schema.xsd";

            // Add the XML schema map to the workbook's XmlMaps collection
            // The Add method returns the index of the newly added map.
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchemaPath);

            // Optionally, set a friendly name for the map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "MyXmlSchemaMap";

            // Save the workbook with the attached XML map
            string outputPath = "output_with_xmlmap.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook loaded from '{excelPath}', XML schema map attached, and saved to '{outputPath}'.");
        }
    }
}