// Title: Remove an unwanted XML map from an Excel workbook by its zero‑based index using Aspose.Cells for .NET
// AI Prompts: Write C# code that deletes an XML map at a given zero‑based index from a workbook and saves the updated file with Aspose.Cells. | Generate a C# snippet that checks the XmlMaps collection count before calling RemoveAt to safely remove an XML map.
// Common Searches: Aspose.Cells C# remove XML map at specific index | How to delete an XML map from an Excel file using Aspose.Cells .NET | Validate XmlMaps collection size before removal in Aspose.Cells | Example of XmlMaps.RemoveAt usage in C# | Remove unwanted XML map from workbook and save changes Aspose.Cells
// Tags: remove xml map Aspose.Cells C# | xmlmaps.removeat method .NET | validate xml map index before removal | delete Excel XML map using Aspose.Cells | save workbook after xml map deletion

using System;
using Aspose.Cells;

// The program loads an existing workbook, checks that the specified zero‑based XML map index is within the XmlMaps collection, removes the XML map with workbook.Worksheets.XmlMaps.RemoveAt, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Index of the XML map to remove (zero‑based)
        int xmlMapIndex = 0; // adjust this value as needed

        // Verify that the index is valid before attempting removal
        if (xmlMapIndex >= 0 && xmlMapIndex < workbook.Worksheets.XmlMaps.Count)
        {
            // Remove the XML map at the specified index
            workbook.Worksheets.XmlMaps.RemoveAt(xmlMapIndex);
        }

        // Save the workbook after the XML map has been removed
        workbook.Save("output.xlsx");
    }
}
