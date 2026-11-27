// Title: Enumerate XmlMapCollection and log each XML map’s Name and RootElementName with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells enumerate all XML maps in a workbook and display their names | How to retrieve root element names of XmlMapCollection using Aspose.Cells .NET | Debugging XML maps in an Excel file with Aspose.Cells C# example | List XmlMap names and root elements from a workbook programmatically | Aspose.Cells .NET code to iterate XmlMapCollection for troubleshooting
// Tags: Aspose.Cells XmlMapCollection iteration | C# retrieve xml map name | Aspose.Cells root element extraction | debugging Excel xml maps .NET | list workbook xml maps programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapDebugger
{
    // The sample loads an existing workbook or creates a new one with a sample XML map, accesses the workbook's XmlMapCollection, iterates through each XmlMap, and writes the map's Name and RootElementName to the console (handling unnamed maps). It then saves the workbook and outputs any errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine the workbook to work with.
                // If a file path is supplied as the first argument and the file exists, load it.
                // Otherwise, create a new workbook and add a sample XML map.
                Workbook workbook;
                if (args.Length > 0 && File.Exists(args[0]))
                {
                    string inputPath = args[0];
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Console.WriteLine("Created a new workbook (no input file supplied).");

                    // Create a temporary XML file to serve as a sample map.
                    // The XmlMapCollection.Add method expects a file path, not a raw XML string.
                    string sampleXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
    <item>
        <Name>Item1</Name>
        <Value>100</Value>
    </item>
    <item>
        <Name>Item2</Name>
        <Value>200</Value>
    </item>
</root>";

                    string tempXmlPath = Path.Combine(Path.GetTempPath(), "SampleMap.xml");
                    File.WriteAllText(tempXmlPath, sampleXml);
                    Console.WriteLine($"Temporary XML file created at \"{tempXmlPath}\".");

                    // Add the XML map to the workbook.
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXmlPath);
                    XmlMap addedMap = workbook.Worksheets.XmlMaps[mapIndex];
                    addedMap.Name = "SampleMap";
                    Console.WriteLine($"Added XML map \"{addedMap.Name}\" with root element \"{addedMap.RootElementName}\".");
                }

                // Retrieve the XmlMapCollection from the workbook.
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Debug output: enumerate each map and log its Name and RootElementName.
                if (xmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps are present in the workbook.");
                }
                else
                {
                    Console.WriteLine($"Found {xmlMaps.Count} XML map(s) in the workbook:");
                    for (int i = 0; i < xmlMaps.Count; i++)
                    {
                        XmlMap map = xmlMaps[i];
                        // Name may be empty if not explicitly set; handle gracefully.
                        string mapName = string.IsNullOrEmpty(map.Name) ? "(unnamed)" : map.Name;
                        string rootName = string.IsNullOrEmpty(map.RootElementName) ? "(no root element)" : map.RootElementName;

                        Console.WriteLine($"  Map #{i + 1}: Name = \"{mapName}\", RootElementName = \"{rootName}\"");
                    }
                }

                // Save the workbook so that the added map (if any) can be inspected later.
                string outputPath = "XmlMapDebugOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Do not hide failures – output full error information.
                Console.Error.WriteLine("An unexpected error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
