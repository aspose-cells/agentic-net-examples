using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportAllXmlMapsDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";

            // Verify the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook containing XML maps
                Workbook workbook = new Workbook(inputPath);

                // Ensure there are XML maps to export
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps found in the workbook.");
                    return;
                }

                // Export each XML map to a separate file
                for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
                {
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[i];
                    string outputPath = $"{xmlMap.Name}_Export.xml";

                    // Export the XML data linked to the current map
                    workbook.ExportXml(xmlMap.Name, outputPath);

                    Console.WriteLine($"Exported XML map '{xmlMap.Name}' to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}