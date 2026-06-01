using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonImportExample
{
    public class JsonImportDemo
    {
        public static void Run()
        {
            try
            {
                // Sample JSON array; each object becomes a row, properties become columns
                string json = @"[
                    { ""Name"": ""Alice"", ""Age"": 30, ""Country"": ""USA"" },
                    { ""Name"": ""Bob"",   ""Age"": 25, ""Country"": ""Canada"" },
                    { ""Name"": ""Charlie"", ""Age"": 28, ""Country"": ""UK"" }
                ]";

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure layout to treat the JSON array as a table
                JsonLayoutOptions options = new JsonLayoutOptions
                {
                    ArrayAsTable = true
                };

                // Import JSON data starting at cell A1 (row 0, column 0)
                JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

                // Save the workbook
                string outputPath = "JsonImportedTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            JsonImportDemo.Run();
        }
    }
}