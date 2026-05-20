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

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook wb = new Workbook(inputPath);

                // Get the collection of XML maps in the workbook
                XmlMapCollection xmlMaps = wb.Worksheets.XmlMaps;

                // Check if there are any XML maps
                if (xmlMaps.Count == 0)
                {
                    Console.WriteLine("No XmlMap found in the workbook.");
                    return;
                }

                // Loop through each XML map and export its XML data
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    XmlMap map = xmlMaps[i];

                    // Define the output file name for the exported XML
                    string outputPath = $"{map.Name}.xml";

                    // Export the XML data linked by the current map
                    wb.ExportXml(map.Name, outputPath);

                    Console.WriteLine($"Exported XML map '{map.Name}' to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}