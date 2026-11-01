// Title: Sort a List of Person objects by Age and merge into an Excel template using Aspose.Cells WorkbookDesigner (C#)
// AI Prompts: Order a collection of C# objects by a numeric property and bind the sorted list to a WorkbookDesigner data source for smart marker processing. | Add a file‑existence check before loading the Excel template and then generate the merged workbook with Aspose.Cells. | Show the correct way to specify the data source name and object in WorkbookDesigner to produce a sorted output in the resulting Excel file.
// Common Searches: C# Aspose.Cells sort list before WorkbookDesigner.Process | How to merge ordered collection into Excel using smart markers in .NET | Check template file existence before using Aspose.Cells WorkbookDesigner | SetDataSource order of parameters Aspose.Cells smart markers example
// Tags: ordered data source for WorkbookDesigner | smart markers template validation Aspose.Cells | C# Aspose.Cells merge sorted list | Excel output from sorted collection using Aspose.Cells | handling FileNotFoundException with Aspose.Cells template

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// The example loads an Excel template, verifies its presence, sorts a List<Person> by Age, assigns the ordered list to the "Persons" data source of a WorkbookDesigner, processes smart markers, and saves the merged result, with exception handling for missing files.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook that contains smart markers
            var workbook = new Workbook(templatePath);

            // Sample data collection
            var persons = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30 },
                new Person { Name = "Bob",     Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            };

            // Sort the collection by Age before merging
            var sortedPersons = persons.OrderBy(p => p.Age).ToList();

            // Use WorkbookDesigner (smart marker processor) to merge data
            var designer = new WorkbookDesigner(workbook);
            // Correct argument order: first the data source name, then the data object
            designer.SetDataSource("Persons", sortedPersons);
            designer.Process();

            // Save the merged workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Result saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
