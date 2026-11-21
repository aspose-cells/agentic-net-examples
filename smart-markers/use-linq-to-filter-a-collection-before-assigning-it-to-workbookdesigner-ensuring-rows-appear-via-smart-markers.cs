// Title: Apply LINQ filtering to a collection and bind the result to WorkbookDesigner for smart‑marker row generation in Aspose.Cells for .NET
// Common Searches: how to apply LINQ filter to a collection before setting WorkbookDesigner data source in Aspose.Cells C# | binding a filtered DataSet to WorkbookDesigner to generate rows with smart markers | populate Excel template smart markers from a LINQ‑filtered List<Person> using Aspose.Cells | filtering data by city and exporting to Excel with smart markers in .NET | Aspose.Cells WorkbookDesigner process after LINQ filtering of source data
// Tags: LINQ filtered collection for WorkbookDesigner | Aspose.Cells smart markers DataSet binding | C# list to DataSet conversion for Excel template | Excel smart marker row generation from filtered data | city-based data filtering for Excel export

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an Excel template containing smart markers, uses LINQ to keep only persons from New York, converts the filtered List<Person> into a DataSet with a 'Persons' DataTable, assigns this DataSet to a WorkbookDesigner, processes the smart markers to expand rows, and saves the populated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Ensure the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the Excel template that contains smart markers (e.g., &=Persons.Name)
            Workbook workbook = new Workbook(templatePath);

            // Sample data collection
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30, City = "New York" },
                new Person { Name = "Bob",     Age = 25, City = "Los Angeles" },
                new Person { Name = "Charlie", Age = 35, City = "New York" },
                new Person { Name = "Diana",   Age = 28, City = "Chicago" }
            };

            // Use LINQ to filter the collection (e.g., only people from New York)
            List<Person> filteredPeople = people
                .Where(p => p.City == "New York")
                .ToList();

            // Convert the filtered list to a DataSet (required by WorkbookDesigner.SetDataSource)
            DataSet dataSet = new DataSet();
            DataTable table = new DataTable("Persons");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Age", typeof(int));
            table.Columns.Add("City", typeof(string));

            foreach (Person p in filteredPeople)
            {
                table.Rows.Add(p.Name, p.Age, p.City);
            }

            dataSet.Tables.Add(table);

            // Assign the DataSet to the WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process(); // Populate rows via smart markers

            // Save the populated workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Simple POCO class representing the data source for smart markers
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string City { get; set; } = string.Empty;
}
