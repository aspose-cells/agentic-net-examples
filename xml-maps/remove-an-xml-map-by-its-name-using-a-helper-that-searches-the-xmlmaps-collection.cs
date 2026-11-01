// Title: Remove a specific XML map by name from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to delete an XML map from an Excel file using Aspose.Cells | how to find and remove an XmlMap by its name in Aspose.Cells .NET | Aspose.Cells remove XmlMap from workbook dynamically | example of removing XML map from .xlsx using Aspose.Cells C# | delete specific XML map in Excel using Aspose.Cells API
// Tags: Aspose.Cells remove XmlMap by name | XML map deletion in Excel workbook C# | dynamic access to XmlMaps collection Aspose | Excel .xlsx XML map manipulation Aspose.Cells | C# helper method find XmlMap by name

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, uses dynamic typing to access its XmlMaps collection, searches for a map with a given name via a helper method, removes the map if present, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string xmlMapName = "MyXmlMap";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Use dynamic to access XmlMaps (available in newer versions of Aspose.Cells)
            dynamic wb = workbook;
            dynamic xmlMaps = wb.XmlMaps;

            // Find the XML map by name
            dynamic mapToRemove = FindXmlMapByName(xmlMaps, xmlMapName);

            if (mapToRemove != null)
            {
                // Remove the map from the collection
                xmlMaps.Remove(mapToRemove);
                Console.WriteLine($"XML map '{xmlMapName}' has been removed.");
            }
            else
            {
                Console.WriteLine($"XML map '{xmlMapName}' was not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method that searches the XmlMaps collection by map name using dynamic typing
    static dynamic FindXmlMapByName(dynamic xmlMaps, string name)
    {
        foreach (dynamic map in xmlMaps)
        {
            if (string.Equals((string)map.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return map;
            }
        }
        return null;
    }
}
