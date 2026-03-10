using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook (any supported format, e.g., .xlsx)
            string inputPath = "input.xlsx";

            // Path to the XML schema (XSD) that defines the XML map
            string xmlSchemaPath = "schema.xsd";

            // Path where the workbook with the added XML map will be saved
            string outputPath = "output_with_xmlmap.xlsx";

            // Load the workbook using the provided lifecycle rule
            Workbook workbook = new Workbook(inputPath);

            // Add an XML map to the workbook (free‑form code, no specific rule exists)
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchemaPath);
            // Optional: retrieve the added map if further processing is needed
            // XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Save the modified workbook using the provided lifecycle rule
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}