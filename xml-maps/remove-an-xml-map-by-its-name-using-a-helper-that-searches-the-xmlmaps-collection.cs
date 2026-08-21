// Title: C# Helper to Remove an XML Map by Name in Aspose.Cells Workbook
// Description: Shows how to locate and delete a specific XmlMap from a Workbook using Aspose.Cells for .NET. The sample adds an XML schema, names the map "EmployeeMap", searches the Worksheets.XmlMaps collection with a case‑insensitive match, removes the map via RemoveAt, and saves the workbook.
// Keywords: Aspose.Cells | C# remove XmlMap | delete XML map .NET | XmlMapCollection RemoveAt | XML map management Aspose | remove XML map by name | Aspose.Cells helper method | search XmlMaps collection | case insensitive map removal | Workbook XML map deletion
// Common Searches: remove xml map Aspose.Cells C# | delete specific XmlMap by name .NET | how to remove XML map from workbook using Aspose | Aspose.Cells XmlMapCollection remove example | C# code to delete XML map in Excel file | search and delete XmlMap in Aspose.Cells
// Developer Intent: Programmatically delete a particular XML map from an Excel workbook based on its name.
// Use Cases: Clean up imported XML maps that are no longer needed, reducing file size. | Automate removal of a map before publishing a workbook to external systems. | Batch‑process a library of workbooks to purge a legacy XML map across all files.
// AI Prompts: Generate a C# method that removes an XmlMap with a given name from an Aspose.Cells Workbook and returns true if the map was found. | Provide code to list all XmlMap names in a workbook and delete those matching a wildcard pattern using Aspose.Cells for .NET. | Explain best practices for safely removing an XmlMap without affecting existing worksheets, data connections, or formulas.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to locate and delete a specific XmlMap from a Workbook using Aspose.Cells for .NET. The sample adds an XML schema, names the map "EmployeeMap", searches the Worksheets.XmlMaps collection with a case‑insensitive match, removes the map via RemoveAt, and saves the workbook.
public class RemoveXmlMapByNameDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add a sample XML map to demonstrate removal
        string xmlSchema = "<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>" +
                           "<xs:element name='Employee'>" +
                           "<xs:complexType><xs:sequence>" +
                           "<xs:element name='ID' type='xs:int'/>" +
                           "<xs:element name='Name' type='xs:string'/>" +
                           "</xs:sequence></xs:complexType></xs:element></xs:schema>";
        int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
        XmlMap map = workbook.Worksheets.XmlMaps[mapIndex];
        map.Name = "EmployeeMap";

        // Remove the XML map by its name using the helper method
        RemoveXmlMapByName(workbook, "EmployeeMap");

        // Determine output path and save the workbook
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RemovedXmlMapDemo.xlsx");
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to: {outputPath}");
    }

    // Helper that searches the XmlMaps collection and removes the map with the specified name
    private static void RemoveXmlMapByName(Workbook workbook, string mapName)
    {
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // Iterate through the collection to find the matching map
        for (int i = 0; i < xmlMaps.Count; i++)
        {
            if (xmlMaps[i].Name.Equals(mapName, StringComparison.OrdinalIgnoreCase))
            {
                // Remove the map at the found index
                xmlMaps.RemoveAt(i);
                // Exit after removal
                break;
            }
        }
    }
}
