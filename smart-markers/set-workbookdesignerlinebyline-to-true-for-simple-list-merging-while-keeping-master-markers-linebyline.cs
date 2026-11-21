// Title: How to merge a simple list into an Excel template line‑by‑line using WorkbookDesigner in Aspose.Cells for C#
// Common Searches: Aspose.Cells C# line by line smart marker merge list example | set WorkbookDesigner.LineByLine true to merge list data into Excel template | merge List<Person> into Excel using smart markers line by line Aspose.Cells | C# WorkbookDesigner linebyline property usage for simple list merging | process smart markers line by line with Aspose.Cells in C#
// Tags: WorkbookDesigner line-by-line merging | Aspose.Cells smart marker list merge | C# merge Excel template with object collection | set WorkbookDesigner data source for smart markers | process smart markers line by line

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLineByLineExample
{
    // Simple data class for demonstration
    // Loads an Excel template containing smart‑marker fields, enables line‑by‑line mode on WorkbookDesigner, assigns a List<Person> as the "Persons" data source, processes the markers, and saves the merged workbook as Result.xlsx.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Load the Excel template that contains merge markers (e.g., &=Persons.Name)
            Workbook workbook = new Workbook("Template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Enable line‑by‑line merging for simple list data
            designer.LineByLine = true;

            // Prepare a simple list of data objects to merge
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John", Age = 30 },
                new Person { Name = "Jane", Age = 25 },
                new Person { Name = "Bob",  Age = 40 }
            };

            // Set the data source for the designer; the name "Persons" matches the master marker in the template
            designer.SetDataSource("Persons", persons);

            // Process the markers and perform the merge
            designer.Process();

            // Save the merged workbook
            workbook.Save("Result.xlsx");
        }
    }
}
