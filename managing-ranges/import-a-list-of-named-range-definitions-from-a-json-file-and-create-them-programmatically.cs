using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

class ImportNamedRanges
{
    // Model matching the JSON structure
    public class NamedRangeDef
    {
        public string Name { get; set; }
        public string RefersTo { get; set; }
    }

    static void Main()
    {
        try
        {
            // Path to the JSON file containing named range definitions
            string jsonPath = "namedRanges.json";

            // Verify the JSON file exists before attempting to read it
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
                return;
            }

            // Read the entire JSON content
            string json = File.ReadAllText(jsonPath);

            // Deserialize JSON into a list of named range definitions
            List<NamedRangeDef> definitions = JsonSerializer.Deserialize<List<NamedRangeDef>>(json);

            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Add each named range to the workbook
            if (definitions != null)
            {
                foreach (var def in definitions)
                {
                    // Add the name to the collection and obtain its index
                    int idx = workbook.Worksheets.Names.Add(def.Name);
                    Name name = workbook.Worksheets.Names[idx];

                    // Ensure the RefersTo formula starts with '=' as required by Aspose.Cells
                    name.RefersTo = def.RefersTo?.StartsWith("=") == true ? def.RefersTo : "=" + def.RefersTo;
                }
            }

            // Optional: sort names for better performance before saving
            workbook.Worksheets.SortNames();

            // Save the workbook (lifecycle save)
            string outputPath = "OutputWithNamedRanges.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}