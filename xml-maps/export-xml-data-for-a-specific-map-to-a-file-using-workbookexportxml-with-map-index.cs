using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportXmlByMapIndexDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "InputWithXmlMaps.xlsx";
            const string outputPath = "ExportedMap.xml";
            const int mapIndex = 0; // change this to the desired map index

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: '{inputPath}'.");
                    return;
                }

                // Load the workbook that contains XML maps
                Workbook workbook = new Workbook(inputPath);

                // Ensure the requested map index exists
                if (workbook.Worksheets.XmlMaps.Count > mapIndex)
                {
                    // Retrieve the XmlMap by index
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                    // Export the XML data using the map's name
                    workbook.ExportXml(xmlMap.Name, outputPath);

                    Console.WriteLine($"XML exported successfully to '{outputPath}' using map index {mapIndex}.");
                }
                else
                {
                    Console.WriteLine($"No XmlMap found at index {mapIndex}. Available maps: {workbook.Worksheets.XmlMaps.Count}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}