// Title: Rename an XML Map in an Aspose.Cells Workbook (C#) and Export the Updated XML
// Description: Loads a workbook (creates a placeholder if missing), verifies the presence of XML maps, renames the first map to a readable identifier, exports the XML using the new map name, and saves the workbook with the updated reference.
// Keywords: Aspose.Cells XML map rename | C# rename XML map | export XML Aspose.Cells | update XML map name workbook | Aspose.Cells XML map manipulation | Excel XML map programming | .NET Aspose.Cells example
// Common Searches: how to rename an XML map in Excel using Aspose.Cells C# | Aspose.Cells export XML after changing map name | programmatically change XML map name in a workbook | C# sample for renaming XML map with Aspose.Cells | check and rename XML maps in an Excel file
// Developer Intent: Programmatically change the name of the first XML map in a workbook and export the XML data using the new map name.
// Use Cases: Create a fallback workbook when the source file is absent, then perform XML map operations. | Validate that a workbook contains at least one XML map before attempting a rename to prevent runtime errors. | Rename an XML map to a more meaningful identifier, export the associated XML, and persist the changes by saving the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that checks for XML maps, renames the first map to a custom name, exports the XML, and saves the workbook while handling missing files. | Write a robust Aspose.Cells example that creates a placeholder workbook if the input file does not exist, then renames an XML map and exports the XML using the new map name.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook (creates a placeholder if missing), verifies the presence of XML maps, renames the first map to a readable identifier, exports the XML using the new map name, and saves the workbook with the updated reference.
    class RenameXmlMapDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";
                const string exportPath = "exported.xml";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook as a placeholder.");
                    // Create a new workbook with a default worksheet
                    var newWb = new Workbook();
                    newWb.Save(inputPath);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure there is at least one XML map
                if (workbook.Worksheets.XmlMaps.Count > 0)
                {
                    // Access the first XML map
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                    // Rename the XML map
                    xmlMap.Name = "ReadableMapName";

                    // Export XML using the new map name to verify the change
                    workbook.ExportXml(xmlMap.Name, exportPath);
                    Console.WriteLine($"XML map renamed and exported to \"{exportPath}\".");
                }
                else
                {
                    Console.WriteLine("No XML maps found in the workbook.");
                }

                // Save the workbook with the updated XML map name
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            RenameXmlMapDemo.Run();
        }
    }
}
