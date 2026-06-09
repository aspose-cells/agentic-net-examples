using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ExportNamedRangeToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Alice");
            worksheet.Cells["B3"].PutValue(25);

            // Define a named range called "People" covering the data
            worksheet.Cells.CreateRange("A1:B3").Name = "People";

            // Retrieve the named range using the GetRangeByName method
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("People");

            // Build a list where each entry contains the cell address and its value
            var cellInfoList = new List<Dictionary<string, string>>();
            foreach (Cell cell in namedRange)
            {
                var entry = new Dictionary<string, string>
                {
                    ["Address"] = cell.Name, // e.g., "A2"
                    ["Value"] = cell.Value?.ToString() ?? string.Empty // Convert value to string
                };
                cellInfoList.Add(entry);
            }

            // Serialize the list to a formatted JSON array
            string jsonResult = JsonSerializer.Serialize(
                cellInfoList,
                new JsonSerializerOptions { WriteIndented = true });

            // Output the JSON
            Console.WriteLine(jsonResult);
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File not found: {ex.FileName}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}