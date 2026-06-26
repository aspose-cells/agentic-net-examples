using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportXmlMapDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "exported.xml";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains XML maps
                Workbook workbook = new Workbook(inputPath);

                // Ensure that at least one XML map is present
                if (workbook.Worksheets.XmlMaps.Count > 0)
                {
                    // Get the first XML map (or select by name/index as needed)
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                    // Export the XML data linked to this map
                    workbook.ExportXml(xmlMap.Name, outputPath);

                    Console.WriteLine($"XML map '{xmlMap.Name}' exported successfully to '{outputPath}'.");
                }
                else
                {
                    Console.WriteLine("No XML maps found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}