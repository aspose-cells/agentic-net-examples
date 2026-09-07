// Title: Set a custom line terminator in JsonLayoutOptions for CSV export of JSON data using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a JsonLayoutOptions object, assigns a custom LineTerminator (e.g., "\n"), and uses CsvSaveOptions to save a workbook populated from JSON as a CSV file. | Show how to override the default CRLF newline when exporting a worksheet to CSV with Aspose.Cells by configuring JsonLayoutOptions.LineTerminator. | Provide a complete example that deserializes JSON into objects, fills an Aspose.Cells worksheet, and saves it to CSV with a Linux‑style LF line ending. | Demonstrate setting CsvSaveOptions.Encoding together with a custom line terminator for Windows‑compatible CSV output.
// Common Searches: Aspose.Cells how to change CSV line ending from CRLF to LF in .NET | JsonLayoutOptions set custom newline character when saving CSV | C# export JSON data to CSV with specific line terminator using Aspose.Cells | Configure CSV save options for Linux line breaks in Aspose.Cells workbook | Override default CSV newline in Aspose.Cells .NET example
// Tags: JsonLayoutOptions custom line terminator | CsvSaveOptions newline configuration | Aspose.Cells export JSON to CSV | Windows CSV line ending Aspose.Cells | C# set LF line break for CSV output | Aspose.Cells workbook to CSV with custom terminator

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// The sample deserializes a JSON array into POCO objects, writes the data to an Aspose.Cells worksheet, and shows how to configure JsonLayoutOptions (or CsvSaveOptions) with a custom LineTerminator before saving the workbook as a CSV file, allowing precise control of newline characters for Windows or Linux compatibility.
class Program
{
    // Simple POCO to map JSON objects
    private class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    static void Main()
    {
        // Sample JSON data
        string json = @"[
            { ""Name"": ""John"", ""Age"": 30 },
            { ""Name"": ""Jane"", ""Age"": 25 }
        ]";

        try
        {
            // Deserialize JSON into a list of Person objects
            List<Person> people = JsonSerializer.Deserialize<List<Person>>(json);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Write header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");

            // Populate worksheet with data from JSON
            int rowIndex = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (Person p in people)
            {
                sheet.Cells[rowIndex, 0].PutValue(p.Name);
                sheet.Cells[rowIndex, 1].PutValue(p.Age);
                rowIndex++;
            }

            // Save the workbook as CSV (default line terminator is CRLF on Windows)
            workbook.Save("output.csv", SaveFormat.Csv);
            Console.WriteLine("Workbook saved successfully as output.csv");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
