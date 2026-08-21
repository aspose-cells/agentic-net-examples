// Title: Aspose.Cells C# – Remove an XML Map by Index with XmlMapCollection.RemoveAt
// Description: Load a workbook, access its XmlMapCollection, validate the target index, call RemoveAt to delete the unwanted XML map, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells remove XML map | XmlMapCollection RemoveAt C# | delete Excel XML map by index | Aspose.Cells XML map removal example | .NET workbook XML map delete
// Common Searches: how to delete an XML map in Aspose.Cells | remove XML map at specific index C# | XmlMapCollection.RemoveAt usage | Aspose.Cells remove unwanted XML map
// Developer Intent: Programmatically delete a specific XML map from an Excel workbook using its zero‑based index.
// Use Cases: Clean up legacy XML maps after data import. | Strip confidential XML mappings before sharing a workbook. | Automate workbook preparation by removing obsolete XML maps.
// AI Prompts: Show a C# snippet that removes an XML map by its name instead of index with Aspose.Cells. | Demonstrate how to enumerate all XML map names, let the user pick one, and safely remove it. | Explain how to verify that an XML map was successfully removed after calling RemoveAt.

using System;
using Aspose.Cells;

// Load a workbook, access its XmlMapCollection, validate the target index, call RemoveAt to delete the unwanted XML map, and save the file using Aspose.Cells for .NET.
class RemoveXmlMapDemo
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of XML maps from the workbook
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // Ensure there is at least one XML map to remove
        if (xmlMaps.Count > 0)
        {
            // Index of the XML map to remove (adjust as needed)
            int indexToRemove = 0;

            // Validate the index before removal
            if (indexToRemove >= 0 && indexToRemove < xmlMaps.Count)
            {
                // Remove the XML map at the specified index
                xmlMaps.RemoveAt(indexToRemove);
                Console.WriteLine($"Removed XML map at index {indexToRemove}.");
            }
            else
            {
                Console.WriteLine("Specified index is out of range.");
            }
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
