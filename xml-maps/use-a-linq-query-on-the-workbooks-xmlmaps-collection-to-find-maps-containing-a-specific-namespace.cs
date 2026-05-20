using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample XML maps (replace with actual file paths as needed)
            // Ensure the schema files exist before adding them
            string schemaPath1 = "schema1.xsd";
            if (File.Exists(schemaPath1))
            {
                try
                {
                    int mapIndex1 = workbook.Worksheets.XmlMaps.Add(schemaPath1);
                    XmlMap xmlMap1 = workbook.Worksheets.XmlMaps[mapIndex1];
                    xmlMap1.Name = "ExampleMap";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add XML map from '{schemaPath1}': {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Schema file not found: {schemaPath1}");
            }

            string schemaPath2 = "schema2.xsd";
            if (File.Exists(schemaPath2))
            {
                try
                {
                    int mapIndex2 = workbook.Worksheets.XmlMaps.Add(schemaPath2);
                    XmlMap xmlMap2 = workbook.Worksheets.XmlMaps[mapIndex2];
                    xmlMap2.Name = "OtherMap";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add XML map from '{schemaPath2}': {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Schema file not found: {schemaPath2}");
            }

            // Namespace to search for within the XML map's DataBinding URL
            string targetNamespace = "example.com";

            // LINQ query: find all XmlMap objects whose DataBinding URL contains the target namespace
            var matchingMaps = workbook.Worksheets.XmlMaps
                .Where(map => map.DataBinding != null &&
                              !string.IsNullOrEmpty(map.DataBinding.Url) &&
                              map.DataBinding.Url.Contains(targetNamespace))
                .ToList();

            // Output the names and URLs of the matching maps
            foreach (var map in matchingMaps)
            {
                Console.WriteLine($"Found map: {map.Name}, URL: {map.DataBinding.Url}");
            }

            // Save the workbook (optional, demonstrates usage of the save rule)
            workbook.Save("XmlMapsQueryResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}