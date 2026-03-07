using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Load the template workbook (XLSX) that contains smart markers.
        var workbook = new Workbook("Template.xlsx");

        // Read JSON data from a file; if the file does not exist, create sample data.
        List<Person> persons;
        const string jsonFile = "persons.json";

        if (File.Exists(jsonFile))
        {
            string json = File.ReadAllText(jsonFile);
            persons = JsonSerializer.Deserialize<List<Person>>(json);
        }
        else
        {
            // Sample data used when the JSON file is missing.
            persons = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 25 }
            };
        }

        // Initialize the WorkbookDesigner and assign the loaded workbook.
        var designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Set the deserialized list as a data source with the name "Person".
        designer.SetDataSource("Person", persons);

        // Process the smart markers in the workbook.
        designer.Process();

        // Save the processed workbook to a new file.
        designer.Workbook.Save("Result.xlsx");
    }
}