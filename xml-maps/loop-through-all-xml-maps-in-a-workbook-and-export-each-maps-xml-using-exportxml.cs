// Title: Export all XML maps from an Excel workbook to individual XML files using Aspose.Cells for .NET
// Common Searches: C# loop through XmlMaps collection in Aspose.Cells and export each map as XML | How to export multiple XML maps from an .xlsx workbook with Aspose.Cells .NET | Using reflection to retrieve XmlMaps property when not directly exposed in Aspose.Cells | Save each XML map to a separate file using Workbook.ExportXml in C# | Example of iterating over workbook XML maps and writing them to disk with Aspose.Cells
// Tags: Aspose.Cells ExportXml multiple maps | iterate XmlMaps collection C# | export workbook XML maps to files | reflection access XmlMaps Aspose.Cells | save each XML map as separate file

using Aspose.Cells;
using System;
using System.Collections;
using System.IO;

// The sample loads an Excel workbook, creates an output folder, uses reflection to obtain the XmlMaps collection, iterates through each map, retrieves its name, and calls Workbook.ExportXml to write each map to a distinct XML file, handling missing files and runtime errors.
class ExportXmlMaps
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the output directory exists
            string outputDir = "ExportedXml";
            Directory.CreateDirectory(outputDir);

            // Use reflection to obtain the XmlMaps collection (avoids compile‑time dependency on the property)
            var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp == null)
            {
                Console.WriteLine("The workbook does not support XML maps in the current Aspose.Cells version.");
                return;
            }

            var xmlMapsObj = xmlMapsProp.GetValue(workbook) as IEnumerable;
            if (xmlMapsObj == null)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Iterate through each XML map and export it
            foreach (var mapObj in xmlMapsObj)
            {
                try
                {
                    var nameProp = mapObj.GetType().GetProperty("Name");
                    string mapName = nameProp?.GetValue(mapObj) as string;

                    if (string.IsNullOrEmpty(mapName))
                        continue;

                    string xmlPath = Path.Combine(outputDir, mapName + ".xml");

                    // Export the XML map to the specified file
                    workbook.ExportXml(mapName, xmlPath);
                    Console.WriteLine($"Exported XML map '{mapName}' to '{xmlPath}'.");
                }
                catch (Exception exMap)
                {
                    Console.WriteLine($"Failed to export an XML map: {exMap.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
