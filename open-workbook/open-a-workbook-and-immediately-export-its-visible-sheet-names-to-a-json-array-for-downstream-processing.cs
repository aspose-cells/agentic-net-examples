// Title: Export visible worksheet names from an Excel workbook to a JSON array using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells get list of visible sheet names and convert to JSON | serialize visible Excel worksheet names to JSON array with Aspose.Cells .NET | retrieve only visible worksheets from a workbook using Aspose.Cells API | export Excel sheet names as JSON using System.Text.Json and Aspose.Cells | Aspose.Cells filter hidden worksheets and output names in JSON
// Tags: Aspose.Cells visible sheet enumeration | C# JSON serialization of worksheet names | Aspose.Cells workbook visibility filter | export Excel sheet list to JSON .NET | list visible worksheets Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, iterates through its worksheets, collects the names of those marked as visible, serializes the name list to a JSON array using System.Text.Json, and writes the JSON string to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Gather names of all visible worksheets
            var visibleSheetNames = new List<string>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use IsVisible property for visibility check (compatible across versions)
                if (sheet.IsVisible)
                {
                    visibleSheetNames.Add(sheet.Name);
                }
            }

            // Convert the list of names to a JSON array string
            string jsonArray = JsonSerializer.Serialize(visibleSheetNames);

            // Output the JSON (could be written to a file or passed downstream)
            Console.WriteLine(jsonArray);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
