using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportXmlMapDataDemo
    {
        public static void Run()
        {
            // Load an existing workbook that contains an XML map.
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure that the workbook has at least one XML map defined.
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps are defined in the workbook.");
                return;
            }

            // Retrieve the first XML map.
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Export the data bound to the XML map.
            workbook.ExportXml(xmlMap.Name, "exportedData.xml");

            Console.WriteLine($"XML data exported successfully to 'exportedData.xml' using map '{xmlMap.Name}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportXmlMapDataDemo.Run();
        }
    }
}