using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonDemo
{
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            string jsonPath = "people.json";

            // If the JSON file does not exist, create a sample one.
            if (!File.Exists(jsonPath))
            {
                var samplePeople = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30, City = "New York" },
                    new Person { Name = "Bob", Age = 25, City = "London" },
                    new Person { Name = "Charlie", Age = 35, City = "Paris" }
                };
                string sampleJson = JsonSerializer.Serialize(samplePeople, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(jsonPath, sampleJson);
            }

            string jsonContent = File.ReadAllText(jsonPath);
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            List<Person> people = JsonSerializer.Deserialize<List<Person>>(jsonContent, options);

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("City");

            for (int i = 0; i < people.Count; i++)
            {
                int rowIndex = i + 1;
                cells[rowIndex, 0].PutValue(people[i].Name);
                cells[rowIndex, 1].PutValue(people[i].Age);
                cells[rowIndex, 2].PutValue(people[i].City);
            }

            string outputPath = "People.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}