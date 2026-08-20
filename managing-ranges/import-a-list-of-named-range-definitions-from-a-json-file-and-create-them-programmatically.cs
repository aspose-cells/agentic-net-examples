// Title: C# – Import Named Ranges from a JSON File into an Aspose.Cells Workbook
// Description: A concise example that reads a JSON file containing named‑range definitions, deserializes them into C# objects, creates a new Aspose.Cells Workbook, adds each range to the Worksheets.Names collection, optionally sorts the names, and saves the workbook.
// Keywords: Aspose.Cells C# named range import | JSON to named range Aspose | C# add named ranges programmatically | deserialize named ranges JSON | Aspose.Cells workbook SortNames | sample code Aspose.Cells JSON | GitHub Aspose.Cells named range example | global C# Excel automation
// Common Searches: how to import named ranges from JSON using Aspose.Cells C# | C# code sample for adding named ranges to an Excel workbook | Aspose.Cells read JSON and create named ranges | sort named ranges in Aspose.Cells workbook | GitHub repository Aspose.Cells named range import example
// Developer Intent: Read a JSON file of named‑range definitions and generate matching named ranges in a new Aspose.Cells workbook using C#.
// Use Cases: Automate the migration of named‑range metadata stored in JSON into Excel files. | Validate the presence and content of a JSON configuration before workbook creation. | Programmatically organize workbook names alphabetically with Worksheets.SortNames() for better maintainability.
// AI Prompts: Write C# code that loads a JSON array of named‑range objects and adds them to an Aspose.Cells workbook, including error handling for missing files. | Show how to set the RefersTo formula for each Name after adding it to the Worksheets.Names collection. | Explain the benefits of calling Worksheets.SortNames() before saving an Aspose.Cells workbook.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsNamedRangeImport
{
    // Model representing a named range definition in the JSON file
    // A concise example that reads a JSON file containing named‑range definitions, deserializes them into C# objects, creates a new Aspose.Cells Workbook, adds each range to the Worksheets.Names collection, optionally sorts the names, and saves the workbook.
    public class NamedRangeDefinition
    {
        public string Name { get; set; }          // The name of the range
        public string RefersTo { get; set; }      // Formula that points to the range, e.g. "=Sheet1!$A$1:$A$10"
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string jsonFilePath = "namedRanges.json";
            const string outputFilePath = "output.xlsx";

            try
            {
                // Verify that the JSON file exists before attempting to read it
                if (!File.Exists(jsonFilePath))
                {
                    Console.WriteLine($"Error: JSON file '{jsonFilePath}' not found.");
                    return;
                }

                // Load and deserialize the JSON content
                string jsonContent = File.ReadAllText(jsonFilePath);
                List<NamedRangeDefinition> namedRanges = JsonSerializer.Deserialize<List<NamedRangeDefinition>>(jsonContent);

                // Guard against null or empty deserialization result
                if (namedRanges == null || namedRanges.Count == 0)
                {
                    Console.WriteLine("No named range definitions found in the JSON file.");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add each named range to the workbook's name collection
                foreach (var rangeDef in namedRanges)
                {
                    // Add the name; the method returns the index of the newly added name
                    int nameIndex = workbook.Worksheets.Names.Add(rangeDef.Name);
                    Name name = workbook.Worksheets.Names[nameIndex];
                    name.RefersTo = rangeDef.RefersTo;
                }

                // Optional: sort the names for better organization before saving
                workbook.Worksheets.SortNames();

                // Save the workbook
                workbook.Save(outputFilePath);
                Console.WriteLine($"Named ranges imported and workbook saved successfully to '{outputFilePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
