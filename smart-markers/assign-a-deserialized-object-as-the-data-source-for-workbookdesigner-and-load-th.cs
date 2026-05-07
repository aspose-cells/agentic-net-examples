using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsDesignerDemo
{
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public List<string> Hobbies { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // Ensure JSON file exists
            string jsonPath = Path.Combine(baseDir, "person.json");
            if (!File.Exists(jsonPath))
            {
                var samplePerson = new Person
                {
                    Name = "John Doe",
                    Age = 30,
                    Hobbies = new List<string> { "Reading", "Traveling", "Swimming" }
                };
                string sampleJson = JsonSerializer.Serialize(samplePerson, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(jsonPath, sampleJson);
            }

            // Read JSON content
            string jsonContent = File.ReadAllText(jsonPath);
            Person deserializedPerson = JsonSerializer.Deserialize<Person>(jsonContent);

            // Ensure template workbook exists
            string templatePath = Path.Combine(baseDir, "Template.xlsx");
            if (!File.Exists(templatePath))
            {
                // Create a simple workbook with a smart marker for demonstration
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Name:");
                ws.Cells["B1"].PutValue("&=Person.Name");
                ws.Cells["A2"].PutValue("Age:");
                ws.Cells["B2"].PutValue("&=Person.Age");
                ws.Cells["A3"].PutValue("Hobbies:");
                ws.Cells["B3"].PutValue("&=Person.Hobbies");
                wb.Save(templatePath);
            }

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set data source
            designer.SetDataSource("Person", deserializedPerson);

            // Process smart markers
            designer.Process();

            // Save result
            string resultPath = Path.Combine(baseDir, "Result.xlsx");
            designer.Workbook.Save(resultPath);

            Console.WriteLine($"Processing complete. Result saved to: {resultPath}");
        }
    }
}