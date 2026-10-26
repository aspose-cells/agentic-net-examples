// Title: How to enable line‑by‑line smart marker merging for a simple list with WorkbookDesigner in Aspose.Cells for .NET
// Common Searches: how to merge a list into an Excel template with WorkbookDesigner line by line | Aspose.Cells set LineByLine true for smart markers | C# example of preserving master markers while merging data | using WorkbookDesigner to process a simple list data source | line-by-line smart marker processing in Aspose.Cells .NET
// Tags: WorkbookDesigner line-by-line mode | Aspose.Cells smart markers list merging | Excel template data source binding C# | preserve master markers Aspose.Cells | set LineByLine property .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;

// This example loads an Excel template, creates a WorkbookDesigner, enables the LineByLine property to merge a simple list of Person objects while keeping master markers, binds the list as the "People" data source, processes the smart markers, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the template workbook that contains master markers
        Workbook workbook = new Workbook("Template.xlsx");

        // Create a WorkbookDesigner for merging data
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Enable line‑by‑line processing for simple list merging while preserving master markers
        designer.LineByLine = true;

        // Example data source: a simple list of objects
        List<Person> people = new List<Person>
        {
            new Person { Name = "John", Age = 30 },
            new Person { Name = "Jane", Age = 25 },
            new Person { Name = "Bob",  Age = 40 }
        };

        // Assign the data source to a name that matches the master marker in the template
        designer.SetDataSource("People", people);

        // Process the markers and merge the data
        designer.Process();

        // Save the merged workbook
        workbook.Save("Result.xlsx");
    }
}

// Simple POCO class used as data source
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}
