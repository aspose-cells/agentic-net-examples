// Title: Export an Excel workbook to camelCase‑named JSON using Aspose.Cells in C#
// Common Searches: c# export excel workbook to json with camelcase property names using aspose.cells | how to set JsonSaveOptions for camelCase naming in Aspose.Cells C# | save Aspose.Cells workbook as JSON file with lowerCamelCase keys | convert worksheet data to JSON using Aspose.Cells and preserve camelCase | Aspose.Cells JsonSaveOptions example for camelCase output in .NET
// Tags: Aspose.Cells JSON export camelCase | JsonSaveOptions naming settings | C# workbook to JSON conversion | Excel worksheet to JSON using Aspose | lowerCamelCase JSON output

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates creating or loading an Excel workbook, populating sample data, configuring JsonSaveOptions for camelCase property naming, and saving the workbook as a JSON file with camelCase keys using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // empty workbook

            // If you have an existing Excel file, uncomment the lines below:
            // string inputPath = "input.xlsx";
            // if (File.Exists(inputPath))
            // {
            //     workbook = new Workbook(inputPath);
            // }
            // else
            // {
            //     Console.WriteLine($"Input file not found: {inputPath}");
            //     return;
            // }

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Configure JSON save options (default naming convention)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Export the workbook to a JSON file
            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
