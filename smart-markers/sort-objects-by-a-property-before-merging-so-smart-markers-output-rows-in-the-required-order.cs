// Title: How to sort a C# collection by a property and merge it into an Excel template using Aspose.Cells smart markers
// Common Searches: C# sort list by age before using Aspose.Cells WorkbookDesigner smart markers | Aspose.Cells setdatasource with ordered DataTable for Excel smart markers | merge sorted collection into Excel template using Aspose.Cells smart markers C# | convert sorted List<Person> to DataTable for Aspose.Cells smart markers | order data for smart markers in Aspose.Cells workbook
// Tags: Aspose.Cells WorkbookDesigner sorted data source | C# sort collection before smart markers | Excel template smart markers ordered rows | convert List to DataTable Aspose.Cells | setdatasource ordered collection Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

// The example loads an Excel template, creates a List<Person>, orders the list by Age, converts the sorted list into a DataTable, adds it to a DataSet, sets this DataSet as the data source for WorkbookDesigner, processes the smart markers, and saves the populated workbook.
public class Person
{
    public string Name { get; set; } = string.Empty; // initialize to avoid nullable warning
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to prevent FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook that contains smart markers
            var workbook = new Workbook(templatePath);

            // Create a collection of objects
            var people = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30 },
                new Person { Name = "Bob",     Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            };

            // Sort the collection by the required property (Age) before merging
            var sortedPeople = people.OrderBy(p => p.Age).ToList();

            // Convert the collection to a DataTable (required by SetDataSource in this API version)
            var dataTable = new DataTable("People");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));

            foreach (var person in sortedPeople)
            {
                dataTable.Rows.Add(person.Name, person.Age);
            }

            var dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            // Process smart markers using WorkbookDesigner
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the resulting workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for debugging
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
