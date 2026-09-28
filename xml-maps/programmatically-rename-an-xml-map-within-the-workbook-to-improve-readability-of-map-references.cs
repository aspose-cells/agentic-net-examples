// Title: Rename an XML map in an Excel workbook with Aspose.Cells for .NET (C# example)
// AI Prompts: Load a workbook, locate the first XmlMap, assign a new descriptive Name, and save the file using Aspose.Cells. | Identify an XmlMap by its original name in a .xlsx file, modify its Name property to a new identifier, and persist the changes with Aspose.Cells.
// Common Searches: C# Aspose.Cells how to change the name of an XML map in an existing .xlsx file | rename xml map in Excel workbook programmatically using Aspose.Cells .NET | set XmlMap.Name property Aspose.Cells example code | update XML map reference after renaming in Aspose.Cells workbook | Aspose.Cells XML map rename without losing data
// Tags: xml map name change Aspose.Cells | set XmlMap.Name property C# | update Excel workbook XML map identifier | modify XML map reference Aspose.Cells | programmatic XML map rename .NET

using System;
using Aspose.Cells;

// The example loads an existing workbook, checks for XML maps, renames the first map to a readable identifier, outputs the new name, and saves the workbook with the updated XML map name.
class RenameXmlMapExample
{
    static void Main()
    {
        // Load an existing workbook that contains an XML map
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure the workbook has at least one XML map
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the first XML map (you can also locate by original name)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Rename the XML map to a more readable name
            xmlMap.Name = "CustomerDataMap";

            // Optionally, you can verify the rename
            Console.WriteLine("XML map renamed to: " + xmlMap.Name);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }

        // Save the workbook with the updated XML map name
        workbook.Save("output.xlsx");
    }
}
