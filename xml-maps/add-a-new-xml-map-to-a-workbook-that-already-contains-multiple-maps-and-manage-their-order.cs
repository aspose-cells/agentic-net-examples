// Title: Add and reorder XML maps in an existing Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: aspnet add additional XML map to existing workbook with Aspose.Cells | c# reorder XmlMapCollection order after adding new map using Aspose.Cells | move newly added XML map to first position in Excel file with Aspose.Cells .NET | clear and rebuild XML map collection to change map order Aspose.Cells C# | how to change XML map sequence in an Excel workbook programmatically with Aspose.Cells
// Tags: add XML map to Excel workbook using Aspose.Cells | reorder XmlMapCollection programmatically | recreate XML map collection order Aspose.Cells | place new XML map at beginning of collection | manage multiple XML maps in Excel file

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace XmlMapOrderDemo
{
    // The C# example loads an existing workbook, ensures it contains at least two XML maps, adds a new XML map from a schema file, displays the current map order, clears the XmlMapCollection, re‑adds the maps in a custom sequence (new map first), verifies the final order, and saves the workbook with the updated XML map arrangement.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to input files – adjust these paths as needed.
            string workbookPath = "input.xlsx";
            string schemaPath1 = "schema1.xsd";
            string schemaPath2 = "schema2.xsd";
            string newSchemaPath = "newSchema.xsd";

            // Validate that required files exist before proceeding.
            if (!System.IO.File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                return;
            }
            if (!System.IO.File.Exists(schemaPath1) ||
                !System.IO.File.Exists(schemaPath2) ||
                !System.IO.File.Exists(newSchemaPath))
            {
                Console.WriteLine("Error: One or more XML schema files are missing.");
                return;
            }

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(workbookPath);
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // -----------------------------------------------------------------
                // Step 1 – Ensure the workbook already has some XML maps.
                // If the collection is empty, add two initial maps for demonstration.
                // -----------------------------------------------------------------
                if (xmlMaps.Count == 0)
                {
                    Console.WriteLine("Workbook has no XML maps. Adding two initial maps.");
                    xmlMaps.Add(schemaPath1);
                    xmlMaps.Add(schemaPath2);
                }

                // -----------------------------------------------------------------
                // Step 2 – Add a new XML map.
                // -----------------------------------------------------------------
                int newMapIndex = xmlMaps.Add(newSchemaPath);
                XmlMap newMap = xmlMaps[newMapIndex];
                Console.WriteLine($"Added new XML map: Index={newMapIndex}, Name=\"{newMap.Name}\"");

                // -----------------------------------------------------------------
                // Step 3 – Display current order of XML maps.
                // -----------------------------------------------------------------
                Console.WriteLine("\nCurrent XML map order:");
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    Console.WriteLine($"  [{i}] Name = \"{xmlMaps[i].Name}\"");
                }

                // -----------------------------------------------------------------
                // Step 4 – Reorder maps: move the newly added map to the first position.
                // Since XmlMapCollection does not expose an Insert method, we rebuild the collection.
                // -----------------------------------------------------------------
                // Capture the schema file paths associated with each map.
                // The XmlMap class does not expose the original schema path, so we maintain our own list.
                // For this demo we know the order we added them, so we construct the list accordingly.
                List<string> orderedSchemaPaths = new List<string>();

                // Existing maps (before adding the new one) were added from schemaPath1 and schemaPath2.
                // The new map was added from newSchemaPath.
                // Desired order: newSchemaPath, schemaPath1, schemaPath2
                orderedSchemaPaths.Add(newSchemaPath);
                orderedSchemaPaths.Add(schemaPath1);
                orderedSchemaPaths.Add(schemaPath2);

                // Clear the existing collection.
                xmlMaps.Clear();

                // Re‑add maps in the new order.
                foreach (string schemaPath in orderedSchemaPaths)
                {
                    int idx = xmlMaps.Add(schemaPath);
                    XmlMap map = xmlMaps[idx];
                    Console.WriteLine($"Re‑added map: Index={idx}, Name=\"{map.Name}\" (Source=\"{schemaPath}\")");
                }

                // -----------------------------------------------------------------
                // Step 5 – Verify final order.
                // -----------------------------------------------------------------
                Console.WriteLine("\nFinal XML map order after reordering:");
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    Console.WriteLine($"  [{i}] Name = \"{xmlMaps[i].Name}\"");
                }

                // Save the workbook with the updated XML map order.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
