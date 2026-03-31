using System;
using Aspose.Cells;

namespace AsposeCellsXmlExport
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file that contains XML mapping
            string excelPath = "input.xlsx";
            // Path where the exported XML will be saved
            string xmlOutputPath = "exported.xml";

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Check if the workbook contains any XML maps
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Retrieve the first XML map
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Export the data bound to the XML map
                workbook.ExportXml(xmlMap.Name, xmlOutputPath);

                Console.WriteLine($"XML exported successfully to '{xmlOutputPath}'.");
            }
            else
            {
                Console.WriteLine("No XML map found in the workbook.");
            }
        }
    }
}