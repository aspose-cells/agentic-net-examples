using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportLinkedXmlData
    {
        public static void Run()
        {
            // Load the workbook that contains XML maps and linked data
            using (Workbook workbook = new Workbook("InputWorkbook.xlsx"))
            {
                // Verify that the workbook has at least one XML map
                if (workbook.Worksheets.XmlMaps.Count > 0)
                {
                    // Retrieve the first XML map (you can select a specific one by index or name)
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                    // Export the XML data linked to this map to an external file
                    // This preserves the relationships defined by the map
                    string outputPath = "ExportedData.xml";
                    workbook.ExportXml(xmlMap.Name, outputPath);

                    Console.WriteLine($"XML data exported successfully to '{outputPath}'.");
                }
                else
                {
                    Console.WriteLine("No XML maps found in the workbook. Export aborted.");
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportLinkedXmlData.Run();
        }
    }
}