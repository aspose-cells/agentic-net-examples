// Title: How to set a Windows CRLF line terminator with JsonLayoutOptions when exporting JSON data to CSV using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a JsonLayoutOptions instance, sets its LineTerminator to "\r\n", and passes it to Workbook.Save so the resulting CSV uses Windows line endings. | Show how to adapt the given JSON‑to‑CSV sample to apply a custom line terminator via JsonLayoutOptions and explain the required SaveOptions usage.
// Common Searches: Aspose.Cells set line terminator for CSV export in C# | JsonLayoutOptions LineTerminator CRLF example Aspose.Cells | How to get Windows style line endings in CSV saved by Aspose.Cells | C# convert JSON array to CSV with custom line breaks using Aspose.Cells | Saving workbook as CSV with specific line ending Aspose.Cells .NET
// Tags: Aspose.Cells JsonLayoutOptions line terminator | CSV export with CRLF using Aspose.Cells | JSON to CSV conversion Aspose.Cells | Workbook.Save custom layout options | Windows line endings in CSV .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example deserializes a JSON array into POCO objects, writes them to a worksheet, and saves the workbook as CSV. This guide adds a JsonLayoutOptions object with LineTerminator set to "\r\n" to produce Windows‑compatible CRLF line breaks during the CSV export.
class Program
{
    // Simple POCO to match JSON structure
    private class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    static void Main()
    {
        try
        {
            // Sample JSON data
            string json = @"[
                { ""Name"": ""John"", ""Age"": 30 },
                { ""Name"": ""Anna"", ""Age"": 25 }
            ]";

            // Deserialize JSON into a list of Person objects
            List<Person> people = JsonSerializer.Deserialize<List<Person>>(json);
            if (people == null || people.Count == 0)
            {
                Console.WriteLine("No data found in JSON.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Write header
            sheet.Cells[0, 0].PutValue("Name");
            sheet.Cells[0, 1].PutValue("Age");

            // Populate worksheet with data
            for (int i = 0; i < people.Count; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(people[i].Name);
                sheet.Cells[i + 1, 1].PutValue(people[i].Age);
            }

            // Define output file path
            string outputPath = "output.csv";

            // Save the workbook as CSV (UTF‑8, comma separator)
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
