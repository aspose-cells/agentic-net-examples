// Title: Insert a new XML map into an existing Excel workbook and set its order using Aspose.Cells for .NET (C#)
// AI Prompts: Add an XML map from a file to a workbook that already contains XML maps and place it at a specific zero‑based index with Aspose.Cells. | Reorder the XmlMapCollection after inserting a new map so the new map appears at the desired position. | Access the XmlMaps property via reflection to maintain compatibility with older Aspose.Cells versions when adding maps.
// Common Searches: c# aspocells add xml map to existing workbook at a specific index | how to change the order of xml maps in an Excel file using Aspose.Cells | using reflection to retrieve XmlMaps collection in Aspose.Cells .NET | insert new xml map into workbook that already has multiple xml maps aspocells
// Tags: add xml map Aspose.Cells | reorder xml map collection Aspose.Cells | reflection access XmlMaps Aspose.Cells | insert xml map at index C# | manage multiple xml maps workbook

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, obtains the XmlMapCollection (using reflection for version safety), adds a new XML map from a file, moves it to a specified zero‑based index, and saves the workbook with the updated map order.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string xmlPath = "NewMapData.xml";
            const string outputPath = "output.xlsx";

            // Verify required files exist
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");
            if (!File.Exists(xmlPath))
                throw new FileNotFoundException($"XML data file not found: {xmlPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Use reflection to access XmlMaps (may not be present in older versions)
            var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp == null)
            {
                Console.WriteLine("The current Aspose.Cells version does not support XML maps.");
            }
            else
            {
                // Get the XmlMapCollection instance
                var xmlMapsObj = xmlMapsProp.GetValue(workbook);
                if (xmlMapsObj == null)
                    throw new InvalidOperationException("Failed to retrieve XmlMaps collection.");

                // Use dynamic to invoke members without compile‑time binding
                dynamic xmlMaps = xmlMapsObj;

                // Add the new XML map
                dynamic newMap = xmlMaps.Add("NewMap", xmlPath);

                // Desired position for the new map (zero‑based index)
                int targetIndex = 1;

                // Reorder only if the target index is within the collection bounds
                if (targetIndex >= 0 && targetIndex < (int)xmlMaps.Count)
                {
                    // Remove the map from its current (last) position
                    xmlMaps.RemoveAt((int)xmlMaps.Count - 1);

                    // Insert the map at the desired index
                    xmlMaps.Insert(targetIndex, newMap);
                }
            }

            // Save the workbook with the updated XML map order
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
