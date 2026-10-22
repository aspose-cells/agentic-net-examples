// Title: Retrieve the root element name of an XML map at a given index with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook using Aspose.Cells, checks the XmlMaps collection size, and prints the RootElementName of the map at a specified index, including try‑catch blocks for unsupported XML mapping. | Show how to safely access workbook.XmlMaps[i].RootElementName in C#, handling cases where the index is out of range or the XmlMaps feature is unavailable in the current Aspose.Cells version.
// Common Searches: aspnet cells get xml map root element name by index | c# aspose.cells retrieve root element of specific xml map | how to check xmlmaps count before reading root element name | error handling for missing xml mapping support in aspose.cells | read excel xml map root element name using aspose.cells c#
// Tags: xml map root element name retrieval Aspose.Cells | access XmlMaps collection by index C# | validate XmlMaps count before usage Aspose.Cells | handle missing XML mapping support Aspose.Cells | read Excel XML map root element C#

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook with Aspose.Cells, confirms that XML mapping is supported, validates the requested map index against the XmlMaps collection count, and outputs the RootElementName of the specified XML map, with comprehensive error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook using dynamic to safely access XmlMaps (may not be available in older versions)
            dynamic workbook = new Workbook(inputPath);

            // Index of the XML map to inspect (adjust as needed)
            int mapIndex = 0;

            // Attempt to retrieve the XmlMap collection
            int mapCount = 0;
            try
            {
                mapCount = workbook.XmlMaps.Count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"XML mapping is not supported in this Aspose.Cells version: {ex.Message}");
                return;
            }

            // Verify that the index is within the collection bounds
            if (mapIndex >= 0 && mapIndex < mapCount)
            {
                // Retrieve the root element name of the specified XML map
                string rootElementName = workbook.XmlMaps[mapIndex].RootElementName;

                // Display the result
                Console.WriteLine($"Root element name of XML map at index {mapIndex}: {rootElementName}");
            }
            else
            {
                Console.WriteLine("Invalid XML map index.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
