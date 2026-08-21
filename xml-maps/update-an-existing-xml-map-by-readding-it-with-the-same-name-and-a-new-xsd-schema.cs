// Title: Replace an existing XML map with a new XSD schema in an Excel workbook using Aspose.Cells for .NET (C#)
// Description: Loads a workbook (or creates one if missing), finds an XmlMap named "MyXmlMap", removes it, adds a new XmlMap from a specified XSD file, restores the original map name, and saves the updated workbook. Includes validation for missing files and basic error handling.
// Keywords: Aspose.Cells XML map replace | C# update Excel XML map | add XmlMap from XSD | remove existing XmlMap | preserve XmlMap name | Excel workbook schema update | Aspose.Cells .NET example
// Common Searches: how to replace an XML map in Excel with Aspose.Cells | Aspose.Cells remove and add XmlMap C# | update Excel XML map schema programmatically | change XSD for existing XmlMap using Aspose | C# code to refresh XML map in workbook
// Developer Intent: Swap an existing XmlMap for a new one based on a different XSD while keeping the original map name unchanged.
// Use Cases: Modernize legacy Excel templates by attaching the latest XSD without breaking downstream integrations. | Automate daily regeneration of reporting workbooks that require an updated XML schema. | Generate a fresh workbook when the source file is absent, then attach the required XML map from a supplied schema.
// AI Prompts: Generate C# code with Aspose.Cells that removes an XmlMap named 'MyXmlMap' and adds a new map from 'newSchema.xsd' while preserving the map name. | Create robust error handling for missing workbook or XSD files when updating XML maps in Aspose.Cells. | Explain how to iterate through XmlMapCollection to locate and delete a specific map before inserting a new one.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapUpdate
{
    // Loads a workbook (or creates one if missing), finds an XmlMap named "MyXmlMap", removes it, adds a new XmlMap from a specified XSD file, restores the original map name, and saves the updated workbook. Includes validation for missing files and basic error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for required files
                string inputPath = "input.xlsx";
                string schemaPath = "newSchema.xsd";
                string outputPath = "output.xlsx";

                // Ensure the input workbook exists; create a new one if missing
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Verify that the XSD schema file exists
                if (!File.Exists(schemaPath))
                {
                    throw new FileNotFoundException($"Schema file \"{schemaPath}\" not found.");
                }

                // Name of the XML map to be updated
                const string mapName = "MyXmlMap";

                // Locate the existing XmlMap by name
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
                int existingIndex = -1;
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    if (xmlMaps[i].Name == mapName)
                    {
                        existingIndex = i;
                        break;
                    }
                }

                // Remove the map if it already exists
                if (existingIndex != -1)
                {
                    xmlMaps.RemoveAt(existingIndex);
                }

                // Add the new XmlMap using the provided XSD schema
                int newIndex = xmlMaps.Add(schemaPath);
                XmlMap newMap = xmlMaps[newIndex];

                // Preserve the original map name
                newMap.Name = mapName;

                // Save the workbook with the updated XML map
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
