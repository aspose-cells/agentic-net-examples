using System;
using Aspose.Cells;

namespace AsposeCellsXmlExportDemo
{
    public class ExportXmlMapData
    {
        public static void Run()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure the workbook contains at least one XML map
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Retrieve the first XML map
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Export the XML data associated with the map to an external file
                workbook.ExportXml(xmlMap.Name, "exportedData.xml");

                Console.WriteLine("XML data exported successfully to 'exportedData.xml'.");
            }
            else
            {
                Console.WriteLine("No XML maps found in the workbook.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportXmlMapData.Run();
        }
    }
}