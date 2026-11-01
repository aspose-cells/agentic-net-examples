// Title: Add a new XML map to an existing Excel workbook and change its order with Aspose.Cells for .NET
// Common Searches: how to insert an XML map at a specific position in an Excel file using Aspose.Cells C# | Aspose.Cells .NET reorder existing XML maps in a workbook | add multiple XML maps to a workbook and set their index with Aspose.Cells | using reflection to access the XmlMaps collection in Aspose.Cells when the property is not public | move a newly added XML map to the second slot in an Excel workbook with Aspose.Cells
// Tags: aspacells add xml map | aspacells reorder xml maps | c# xmlmaps collection insert | aspacells reflection xmlmaps | excel xml map ordering with aspacells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, reads an XML file, uses reflection to obtain the XmlMaps collection, adds a new XML map, repositions it to a desired index (e.g., second), and saves the workbook, demonstrating how to manage multiple XML maps and their order with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string xmlPath = "newmap.xml";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Workbook file not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the XML map source file exists
            if (!File.Exists(xmlPath))
                throw new FileNotFoundException($"XML map file not found: {xmlPath}");

            // Read the XML content that will be used for the new map
            string xmlContent = File.ReadAllText(xmlPath);
            string xmlMapName = "NewMap";

            // Use reflection to access XmlMaps (may not be available in older versions)
            var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp == null)
                throw new NotSupportedException("The current Aspose.Cells version does not support XML maps.");

            // Get the XmlMapCollection instance
            object xmlMapsObj = xmlMapsProp.GetValue(workbook);
            dynamic xmlMaps = xmlMapsObj; // use dynamic to invoke members

            // Add the new XML map; the method returns the index of the added map
            int addedMapIndex = xmlMaps.Add(xmlMapName, xmlContent);

            // Desired position for the new map (e.g., make it the second map in the collection)
            int desiredIndex = 1; // zero‑based index

            // Reorder the XML maps if needed
            if (addedMapIndex != desiredIndex &&
                addedMapIndex >= 0 &&
                desiredIndex >= 0 &&
                desiredIndex < (int)xmlMaps.Count)
            {
                // Retrieve the added map instance
                var addedMap = xmlMaps[addedMapIndex];

                // Remove it from its current position
                xmlMaps.RemoveAt(addedMapIndex);

                // Insert it at the desired position
                xmlMaps.Insert(desiredIndex, addedMap);
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
