using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportXmlDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the Excel file that contains an XML map
            string inputFileName = "InputWithXmlMap.xlsx";
            string inputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, inputFileName);

            // Path where the exported XML will be saved
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExportedData.xml");

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook has at least one XML map
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
}