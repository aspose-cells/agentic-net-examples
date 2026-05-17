using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportAllXmlMaps
    {
        public static void Run()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Get the collection of XML maps in the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Check if any XML maps are present
            if (xmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Loop through each XML map and export its data to a separate XML file
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                XmlMap map = xmlMaps[i];

                // Create a safe file name using the map name (fallback to index if name is empty)
                string safeMapName = string.IsNullOrWhiteSpace(map.Name) ? $"Map{i}" : map.Name;
                // Replace any invalid filename characters
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    safeMapName = safeMapName.Replace(c, '_');
                }

                string outputPath = $"{safeMapName}.xml";

                // Export the XML data for the current map
                workbook.ExportXml(map.Name, outputPath);

                Console.WriteLine($"Exported XML map '{map.Name}' to file: {outputPath}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportAllXmlMaps.Run();
        }
    }
}