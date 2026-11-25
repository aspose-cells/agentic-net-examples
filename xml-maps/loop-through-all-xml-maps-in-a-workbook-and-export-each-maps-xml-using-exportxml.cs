// Title: Export every XML map from an Excel workbook to separate .xml files using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells loop through XmlMapCollection and export each map to a file | How to save multiple XML maps from an Excel workbook with Aspose.Cells | ExportXml for all XML maps in a workbook using Aspose.Cells .NET | Create individual XML files for each XML map in an Excel file with Aspose.Cells | Batch export of XML maps from Excel using Aspose.Cells C# example
// Tags: Aspose.Cells ExportXml multiple maps | C# iterate XmlMapCollection | export XML maps from Excel workbook | generate separate XML files per map | batch XML map export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExportAllXmlMaps
{
    // The sample loads an Excel workbook, checks for XML maps, and iterates through the XmlMapCollection. For each map it determines a name, builds an output path in a user‑specified folder, and calls Workbook.ExportXml to write the map's XML data to its own .xml file, handling errors and creating the output directory when needed.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and output directory
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: ExportAllXmlMaps <inputWorkbookPath> <outputDirectory>");
                return;
            }

            string inputPath = args[0];
            string outputDir = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The workbook file '{inputPath}' does not exist.");
                return;
            }

            if (!Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                    Console.WriteLine($"Created output directory: {outputDir}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Unable to create output directory '{outputDir}'. {ex.Message}");
                    return;
                }
            }

            try
            {
                // Load the workbook
                using (Workbook workbook = new Workbook(inputPath))
                {
                    XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                    if (xmlMaps.Count == 0)
                    {
                        Console.WriteLine("No XML maps found in the workbook.");
                        return;
                    }

                    Console.WriteLine($"Found {xmlMaps.Count} XML map(s) in the workbook.");

                    for (int i = 0; i < xmlMaps.Count; i++)
                    {
                        XmlMap map = xmlMaps[i];
                        // Ensure the map has a valid name; if not, generate one
                        string mapName = string.IsNullOrWhiteSpace(map.Name) ? $"Map_{i}" : map.Name;

                        // Build the output file path
                        string outputPath = Path.Combine(outputDir, $"{mapName}.xml");

                        try
                        {
                            // Export the XML data for the current map
                            workbook.ExportXml(mapName, outputPath);
                            Console.WriteLine($"Exported XML map '{mapName}' to '{outputPath}'.");
                        }
                        catch (Exception exportEx)
                        {
                            Console.WriteLine($"Failed to export map '{mapName}'. {exportEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
