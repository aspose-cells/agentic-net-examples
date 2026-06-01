using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Specify the name of the XML map to be removed
        string mapNameToRemove = "Employee";

        // Remove the XML map using the helper method
        RemoveXmlMapByName(workbook, mapNameToRemove);

        // Save the workbook after the XML map has been removed
        workbook.Save("output.xlsx");
    }

    // Helper that searches the XmlMaps collection and removes the map with the given name
    static void RemoveXmlMapByName(Workbook workbook, string mapName)
    {
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // Iterate through the collection to locate the map with the specified name
        for (int i = 0; i < xmlMaps.Count; i++)
        {
            XmlMap map = xmlMaps[i];
            if (map.Name == mapName)
            {
                // Remove the map at the found index
                xmlMaps.RemoveAt(i);
                // Exit after removal since the map has been found
                break;
            }
        }
    }
}