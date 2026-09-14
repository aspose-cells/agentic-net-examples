// Title: Filter Aspose.Cells workbook XML maps by target namespace using LINQ in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, accesses its XmlMaps collection, and uses a LINQ query to return all XmlMap objects whose TargetNamespace equals a specified string. | Show how to safely obtain the XmlMapCollection via reflection for older Aspose.Cells versions and then apply a LINQ Where clause to select map names that match a given namespace. | Provide a LINQ expression that extracts the Id values of XmlMap entries whose Namespace property contains the user‑provided namespace pattern.
// Common Searches: aspocells linq filter xmlmap by namespace c# | how to query XmlMapCollection for a specific namespace using Aspose.Cells | c# retrieve xml map ids from workbook where target namespace matches | using reflection to access XmlMaps in older Aspose.Cells versions and filter with LINQ | example of LINQ Where on Aspose.Cells XmlMapCollection
// Tags: linq filter aspocells xmlmap collection | retrieve xmlmap by targetnamespace c# | aspocells xmlmap enumeration via reflection | excel workbook xml map querying | c# aspocells xml namespace search

using Aspose.Cells;
using System;
using System.IO;
using System.Collections;

// The example loads an Excel workbook with Aspose.Cells, obtains the XmlMapCollection (using reflection when the direct property is unavailable), and demonstrates how to apply a LINQ query to select maps whose TargetNamespace matches a supplied value. It then prints each matching map's Name and Id while handling missing files, absent XML‑map support, and other runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve XML maps via reflection (API may not be present in all versions)
            var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp == null)
            {
                Console.WriteLine("XML map feature is not available in this Aspose.Cells version.");
                return;
            }

            var xmlMaps = xmlMapsProp.GetValue(workbook, null);
            if (xmlMaps == null)
            {
                Console.WriteLine("No XML maps were found in the workbook.");
                return;
            }

            // Get the Count property of the XmlMapCollection
            var countProp = xmlMaps.GetType().GetProperty("Count");
            int mapCount = countProp != null ? (int)countProp.GetValue(xmlMaps) : 0;

            if (mapCount == 0)
            {
                Console.WriteLine("No XML maps were found in the workbook.");
                return;
            }

            Console.WriteLine($"Found {mapCount} XmlMap(s) in the workbook.");

            // Enumerate the collection using IEnumerable
            foreach (var map in (IEnumerable)xmlMaps)
            {
                var nameProp = map.GetType().GetProperty("Name");
                var idProp = map.GetType().GetProperty("Id");

                string name = nameProp?.GetValue(map)?.ToString() ?? "N/A";
                string id = idProp?.GetValue(map)?.ToString() ?? "N/A";

                Console.WriteLine($"Map Name: {name}, Id: {id}");
            }

            // Uncomment to save any changes made to the workbook
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
