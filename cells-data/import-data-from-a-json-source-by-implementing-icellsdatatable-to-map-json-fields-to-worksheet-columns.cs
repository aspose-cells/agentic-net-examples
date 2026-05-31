using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsDemo
{
    public class JsonToExcelImporter
    {
        // Model representing each JSON object
        public class Person
        {
            public string Name { get; set; } = string.Empty;
            public string DisplayName { get; set; } = string.Empty;
            public bool ShowDisplayName { get; set; }
        }

        public static void Run()
        {
            try
            {
                // Sample JSON array
                string json = "[{\"Name\":\"Name\",\"DisplayName\":\"DisplayName\",\"ShowDisplayName\":false}," +
                              "{\"Name\":\"Nameone\",\"DisplayName\":\"DisplayNameone\",\"ShowDisplayName\":true}," +
                              "{\"Name\":\"Nametwo\",\"DisplayName\":\"DisplayNametwo\",\"ShowDisplayName\":false}]";

                // Deserialize JSON into a list of Person objects
                List<Person>? persons = JsonSerializer.Deserialize<List<Person>>(json);
                if (persons == null)
                {
                    Console.WriteLine("Failed to deserialize JSON.");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Build ICellsDataTable from the list using the factory
                ICellsDataTable dataTable = workbook.CellsDataTableFactory.GetInstance(persons);

                // Import the data table into the worksheet starting at cell A1 (row 0, column 0)
                sheet.Cells.ImportData(dataTable, 0, 0, new ImportTableOptions());

                // Define output file path
                string outputPath = "JsonImportOutput.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            JsonToExcelImporter.Run();
        }
    }
}