using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportXmlFromWorkbook
    {
        public static void Run()
        {
            // Path to the existing Excel file that contains an XML map
            string inputPath = "InputWorkbook.xlsx";

            // Path where the exported XML will be saved
            string outputPath = "ExportedData.xml";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure that the workbook has at least one XML map defined
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Retrieve the first XML map
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Export the XML data linked to the map
                workbook.ExportXml(xmlMap.Name, outputPath);

                Console.WriteLine($"XML exported successfully to '{outputPath}'.");
            }
            else
            {
                Console.WriteLine("No XML map found in the workbook.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportXmlFromWorkbook.Run();
        }
    }
}