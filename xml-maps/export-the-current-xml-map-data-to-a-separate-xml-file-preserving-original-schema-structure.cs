using System;
using Aspose.Cells;

namespace AsposeCellsExportXmlExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains an XML map
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Ensure that at least one XML map is present
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Retrieve the first XML map
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Export the XML data linked by this map to a separate file
                // This preserves the original schema structure defined in the map
                workbook.ExportXml(xmlMap.Name, "ExportedData.xml");

                Console.WriteLine("XML data exported successfully to ExportedData.xml");
            }
            else
            {
                Console.WriteLine("No XML maps found in the workbook.");
            }
        }
    }
}