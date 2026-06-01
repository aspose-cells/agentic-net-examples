using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapUpdate
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths
                string inputWorkbookPath = "input.xlsx";
                string newSchemaPath = "newSchema.xsd";
                string outputWorkbookPath = "output.xlsx";

                // Verify input files exist
                if (!File.Exists(inputWorkbookPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputWorkbookPath}");
                    return;
                }

                if (!File.Exists(newSchemaPath))
                {
                    Console.WriteLine($"XSD schema file not found: {newSchemaPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Name of the XML map to replace
                string targetMapName = "MyMap";

                // Access XML maps collection
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Find existing map index
                int existingMapIndex = -1;
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    if (xmlMaps[i].Name == targetMapName)
                    {
                        existingMapIndex = i;
                        break;
                    }
                }

                // Remove existing map if found
                if (existingMapIndex != -1)
                {
                    xmlMaps.RemoveAt(existingMapIndex);
                }

                // Add new XML map from the XSD schema
                int newMapIndex = xmlMaps.Add(newSchemaPath);

                // Preserve original map name
                xmlMaps[newMapIndex].Name = targetMapName;

                // Save the updated workbook
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved successfully to {outputWorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}