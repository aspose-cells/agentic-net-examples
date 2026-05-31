using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
    public string Department { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Sample data collection
        List<Person> persons = new List<Person>
        {
            new Person { Name = "John",  Age = 28, Department = "Sales" },
            new Person { Name = "Alice", Age = 35, Department = "HR" },
            new Person { Name = "Bob",   Age = 42, Department = "IT" },
            new Person { Name = "Eve",   Age = 31, Department = "Finance" }
        };

        // LINQ filter: only persons older than 30
        var filteredPersons = persons.Where(p => p.Age > 30).ToList();

        // Create a new workbook and set up smart markers
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("Department");

        // Smart marker row (will be repeated for each data item)
        sheet.Cells["A2"].PutValue("&Person.Name");
        sheet.Cells["B2"].PutValue("&Person.Age");
        sheet.Cells["C2"].PutValue("&Person.Department");

        // Define the range that contains the smart markers (required when LineByLine = false)
        sheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

        // Initialize WorkbookDesigner with the workbook
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            LineByLine = false // use range smart markers instead of line‑by‑line processing
        };

        // Bind the filtered collection to the smart marker name "Person"
        designer.SetDataSource("Person", filteredPersons);

        // Process the smart markers and populate the worksheet
        designer.Process();

        // Save the resulting workbook
        workbook.Save("FilteredPersons.xlsx");
    }
}