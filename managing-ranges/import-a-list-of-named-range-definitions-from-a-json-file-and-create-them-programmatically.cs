// Title: Import Named Ranges from JSON into an Aspose.Cells Workbook (C#)
// Description: C# sample that reads a JSON file containing named‑range definitions, deserializes them into objects, adds each name to a workbook’s Names collection, sets the RefersTo formula, optionally sorts the names, and saves the workbook as XLSX.
// Keywords: Aspose.Cells | C# | JSON | named ranges | import named ranges | create named ranges programmatically | add workbook names | RefersTo formula | sort named ranges | Excel automation | template workbook
// Common Searches: how to import named ranges from JSON using Aspose.Cells C# | C# code to add named ranges to an Excel workbook | Aspose.Cells read JSON and create named ranges | sort named ranges Aspose.Cells | update workbook names from configuration file
// Developer Intent: Add or update named ranges in an Excel workbook based on JSON‑defined specifications.
// Use Cases: Load a base template and inject business‑specific named ranges defined in a JSON config before generating a report. | Consume a service that returns range definitions in JSON and programmatically apply them to a workbook for downstream calculations. | Synchronize named ranges across multiple workbooks by reading a shared JSON definition file and applying the changes in batch.
// AI Prompts: Generate C# code with Aspose.Cells that reads a JSON array of named‑range objects and adds each as a workbook name with the correct RefersTo formula. | Enhance the sample with validation for the RefersTo string, duplicate‑name handling, and detailed error messages. | Show how to modify existing named ranges when the JSON file contains updated definitions while preserving ranges not present in the file.

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// C# sample that reads a JSON file containing named‑range definitions, deserializes them into objects, adds each name to a workbook’s Names collection, sets the RefersTo formula, optionally sorts the names, and saves the workbook as XLSX.
class ImportNamedRanges
{
    // Represents a named range definition in the JSON file
    public class NamedRange
    {
        public string Name { get; set; }          // The name of the range
        public string RefersTo { get; set; }      // The formula reference, e.g. "=Sheet1!$A$1:$A$10"
    }

    static void Main()
    {
        try
        {
            // ---------- Create / Load ----------
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook(); // create new workbook

            // ---------- Read JSON ----------
            string jsonPath = "namedRanges.json";

            if (File.Exists(jsonPath))
            {
                // Read the entire JSON file as a string
                string jsonContent = File.ReadAllText(jsonPath);

                // Deserialize JSON into a list of NamedRange objects
                List<NamedRange> namedRanges = JsonSerializer.Deserialize<List<NamedRange>>(jsonContent);

                // ---------- Create Named Ranges ----------
                if (namedRanges != null)
                {
                    foreach (var nr in namedRanges)
                    {
                        // Add the name to the workbook's name collection; Add returns the index of the new name
                        int index = workbook.Worksheets.Names.Add(nr.Name);

                        // Set the reference for the named range (must start with '=')
                        workbook.Worksheets.Names[index].RefersTo = nr.RefersTo;
                    }

                    // Optional: sort the names for better organization before saving
                    workbook.Worksheets.SortNames();
                }
            }
            else
            {
                Console.WriteLine($"JSON file '{jsonPath}' not found. No named ranges will be added.");
            }

            // ---------- Save ----------
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
