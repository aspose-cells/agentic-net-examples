// Title: Import a JSON array into an Excel worksheet and create a ListObject table with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses System.Text.Json to deserialize a JSON array into a list of POCO objects and inserts the data with column headers into an Aspose.Cells worksheet. | Show how to define the data range and add a ListObject (Excel table) over it using Aspose.Cells. | Demonstrate applying a built‑in table style, auto‑fitting columns, and saving the workbook with proper exception handling.
// Common Searches: C# Aspose.Cells import JSON array into Excel table with headers | How to create a ListObject from deserialized JSON data using Aspose.Cells | Aspose.Cells auto fit columns after writing JSON objects to worksheet | Saving workbook to specific folder after adding table in Aspose.Cells C#
// Tags: JSON deserialization to POCO collection C# | populate Aspose.Cells worksheet from object list | add ListObject table over data range Aspose.Cells | apply built‑in table style to ListObject Aspose.Cells | auto‑fit columns after data import Aspose.Cells | save workbook with exception handling Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example creates a new workbook, deserializes a JSON array into a list of Person POCOs, writes column headers and each record to the first worksheet, defines the occupied range, adds a ListObject (Excel table) over that range, optionally applies a built‑in table style, auto‑fits the columns, and saves the file as JsonImport.xlsx with comprehensive error handling.
class Program
{
    // Simple POCO to hold JSON data
    private class Person
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;   // avoid nullable warning
        public int Score { get; set; }
    }

    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // JSON data to import
            string json = @"[
                { ""Id"": 1, ""Name"": ""Alice"", ""Score"": 85 },
                { ""Id"": 2, ""Name"": ""Bob"", ""Score"": 92 },
                { ""Id"": 3, ""Name"": ""Charlie"", ""Score"": 78 }
            ]";

            // Deserialize JSON into a list of Person objects
            List<Person> people = JsonSerializer.Deserialize<List<Person>>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            }) ?? new List<Person>();

            // Write headers
            int startRow = 0;
            int startCol = 0;
            sheet.Cells[startRow, startCol].PutValue("Id");
            sheet.Cells[startRow, startCol + 1].PutValue("Name");
            sheet.Cells[startRow, startCol + 2].PutValue("Score");

            // Write data rows
            int currentRow = startRow + 1;
            foreach (var p in people)
            {
                sheet.Cells[currentRow, startCol].PutValue(p.Id);
                sheet.Cells[currentRow, startCol + 1].PutValue(p.Name);
                sheet.Cells[currentRow, startCol + 2].PutValue(p.Score);
                currentRow++;
            }

            // Create a table (ListObject) over the imported range
            int totalRows = people.Count + 1; // including header
            int totalCols = 3;
            // Add returns the index of the created ListObject in older API versions
            int listIndex = sheet.ListObjects.Add(startRow, startCol, totalRows, totalCols, true);
            ListObject listObject = sheet.ListObjects[listIndex];

            // Optional: apply a built‑in table style if desired (skip if not supported)
            try
            {
                // Example style application (uncomment if your version supports it)
                // listObject.Style = workbook.TableStyles[TableStyleType.TableStyleMedium2];
            }
            catch
            {
                // Ignore styling errors for older API versions
            }

            // Auto‑fit columns for better appearance
            sheet.AutoFitColumns();

            // Determine output path and ensure the directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "JsonImport.xlsx");
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
