// Title: Enumerate XML maps in an Excel workbook and print each map name using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# console application that opens a .xlsx file with Aspose.Cells, accesses the Workbook.XmlMaps collection, and writes the Name of every XmlMap to the console, including error handling for missing files. | Write a .NET code snippet that uses reflection to safely obtain the XmlMaps property from a Workbook object and iterates through the collection to display each XML map's name.
// Common Searches: Aspose.Cells C# list all XML maps in an existing workbook | How to read XmlMap names from an Excel file using Aspose.Cells .NET | C# example for retrieving Workbook.XmlMaps collection with Aspose.Cells | Print XML map names to console with Aspose.Cells for .NET | Fallback method for accessing XmlMaps property in older Aspose.Cells versions
// Tags: xml map enumeration Aspose.Cells .NET | Workbook.XmlMaps access C# | console output xml map names | reflection based XmlMaps retrieval | handle missing XmlMaps property Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Collections;

// // Loads an Excel workbook, safely obtains the XmlMaps collection (using reflection when necessary), iterates through each XmlMap, extracts its Name, and writes the names to the console while handling missing files and absent XmlMaps support.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Attempt to get the XmlMaps property via reflection (covers versions where XmlMaps may be absent)
            var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp == null)
            {
                Console.WriteLine("The XmlMaps property is not available in this Aspose.Cells version.");
                return;
            }

            // Retrieve the collection of XML maps
            var xmlMaps = xmlMapsProp.GetValue(workbook, null) as IEnumerable;
            if (xmlMaps == null)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Iterate through each XmlMap and output its name
            foreach (var map in xmlMaps)
            {
                var nameProp = map.GetType().GetProperty("Name");
                string mapName = nameProp?.GetValue(map, null)?.ToString() ?? "<Unnamed>";
                Console.WriteLine(mapName);
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
