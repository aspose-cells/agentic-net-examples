// Title: Get the root element name of an XML map by index using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, verifies XML map support, and prints the RootElementName of the map at a specified zero‑based position. | Show how to safely obtain the XmlMapCollection from a Workbook object and catch RuntimeBinderException when the Aspose.Cells version does not include XML map functionality. | Demonstrate validating the requested index against workbook.XmlMaps.Count before accessing the selected XmlMap's RootElementName.
// Common Searches: aspocells c# retrieve xml map root element name by index | how to check if a workbook supports xml maps in Aspose.Cells | c# validate xml map index before getting RootElementName with Aspose.Cells | aspocells get xml map collection count in C# | c# handle missing xml map feature in older Aspose.Cells versions
// Tags: aspocells xmlmap rootelementname extraction | c# workbook xmlmaps index bounds check | aspocells xmlmap collection version check | c# handle missing xmlmap support in Aspose.Cells | aspocells retrieve xml map root element by position

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the input Excel file exists, loads it into an Aspose.Cells Workbook, attempts to access the XmlMaps collection while handling versions that lack XML map support, validates the provided map index against the collection count, retrieves the RootElementName of the XML map at that index, and prints the result; any errors are caught and reported.
class Program
{
    static void Main()
    {
        try
        {
            const string filePath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook using dynamic to safely access members that may not exist in older versions
            dynamic workbook = new Workbook(filePath);

            // Attempt to retrieve the XmlMapCollection; if not supported, handle gracefully
            dynamic xmlMaps = null;
            try
            {
                xmlMaps = workbook.XmlMaps;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("The loaded Aspose.Cells version does not support XML maps.");
                return;
            }

            // Index of the XML map to retrieve (adjust as needed)
            int mapIndex = 0;

            // Validate the index against the collection count
            if (mapIndex >= 0 && mapIndex < (int)xmlMaps.Count)
            {
                // Get the root element name of the selected XML map
                string rootElementName = xmlMaps[mapIndex].RootElementName;

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
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
