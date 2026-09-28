// Title: Export all XML maps from an Excel workbook to separate XML files using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# console application that opens a specified .xlsx file with Aspose.Cells, enumerates every XmlMap in the workbook, and writes each map to its own .xml file in a designated output folder. | Provide .NET code that uses reflection to obtain the Workbook.XmlMaps collection for compatibility with different Aspose.Cells versions, then calls Workbook.ExportXml for each map name. | Create a snippet that creates an output directory, logs the success or failure of exporting each XML map, and handles missing input file or absent XmlMaps gracefully.
// Common Searches: how to loop through XmlMaps in a workbook with Aspose.Cells C# | export each XML map to a separate file using Aspose.Cells .NET | using reflection to access XmlMaps property in older Aspose.Cells versions | save Excel XML maps as individual .xml files programmatically | Aspose.Cells ExportXml example for multiple maps
// Tags: export xml maps Aspose.Cells C# | iterate workbook XmlMaps collection | Workbook.ExportXml per map | reflection fallback for Aspose.Cells XmlMaps | create output folder for exported xml

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The sample loads an input.xlsx workbook, verifies the presence of the XmlMaps collection via reflection, creates an 'XmlMapsOutput' directory, iterates each XML map, and calls Workbook.ExportXml to write each map to a separate .xml file while logging successes and handling errors.
class ExportXmlMaps
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the output directory exists
            string outputDir = "XmlMapsOutput";
            Directory.CreateDirectory(outputDir);

            // Use reflection to obtain the XmlMaps collection (covers API variations)
            PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps", BindingFlags.Public | BindingFlags.Instance);
            if (xmlMapsProp == null)
            {
                Console.WriteLine("The loaded Aspose.Cells version does not expose XmlMaps.");
                return;
            }

            var xmlMaps = xmlMapsProp.GetValue(workbook) as System.Collections.IEnumerable;
            if (xmlMaps == null)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Iterate through each XML map and export it
            foreach (object mapObj in xmlMaps)
            {
                // Use reflection to read the map's Name property
                PropertyInfo nameProp = mapObj.GetType().GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
                if (nameProp == null) continue;

                string mapName = nameProp.GetValue(mapObj) as string;
                if (string.IsNullOrEmpty(mapName)) continue;

                string outputPath = Path.Combine(outputDir, mapName + ".xml");

                try
                {
                    // Export the XML map to the specified file
                    workbook.ExportXml(mapName, outputPath);
                    Console.WriteLine($"Exported XML map '{mapName}' to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export map '{mapName}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
