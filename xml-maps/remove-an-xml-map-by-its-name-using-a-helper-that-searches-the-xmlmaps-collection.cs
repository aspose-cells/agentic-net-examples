// Title: How to delete a specific XML map by its name from an Excel workbook using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that locates an XML map with a given name in a workbook and removes it from the Worksheets.XmlMaps collection. | Create a reusable helper method that returns the index of a named XML map in a Workbook, then use it to delete the map and save the updated file.
// Common Searches: asp.net remove xml map from workbook using aspose.cells c# | c# find and delete xml map by name in excel file with aspose | how to get index of a specific xml map in aspose.cells workbook | remove unwanted xml map from worksheets.xmlmaps collection | aspose.cells delete xml map before saving workbook
// Tags: delete xml map Aspose.Cells C# | xml map index lookup Aspose.Cells | workbook xml maps collection manipulation | remove xml map from Excel file using Aspose | c# helper method find xml map index

using Aspose.Cells;
using System;

// The example loads an Excel workbook, uses a helper to locate the index of an XML map named "MyXmlMap" in the Worksheets.XmlMaps collection, removes the map if it exists, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (uses the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Name of the XML map to be removed
        string xmlMapName = "MyXmlMap";

        // Find the index of the XML map using the helper
        int mapIndex = FindXmlMapIndex(workbook, xmlMapName);

        // If the map exists, remove it from the collection
        if (mapIndex >= 0)
        {
            workbook.Worksheets.XmlMaps.RemoveAt(mapIndex);
        }

        // Save the workbook (uses the provided save rule)
        workbook.Save("output.xlsx");
    }

    // Helper that searches the XmlMaps collection for a map with the specified name
    static int FindXmlMapIndex(Workbook workbook, string mapName)
    {
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
        for (int i = 0; i < xmlMaps.Count; i++)
        {
            if (xmlMaps[i].Name.Equals(mapName, StringComparison.OrdinalIgnoreCase))
            {
                return i; // Return the index of the matching XML map
            }
        }
        return -1; // Not found
    }
}
