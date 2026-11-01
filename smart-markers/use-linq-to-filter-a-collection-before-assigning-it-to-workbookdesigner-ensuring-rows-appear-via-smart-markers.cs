// Title: Filter a List<Person> with LINQ and bind the result to an Aspose.Cells smart marker using WorkbookDesigner in C#
// AI Prompts: Use a LINQ Where clause (e.g., Age > 30) on a List<Person> and pass the filtered collection to WorkbookDesigner.SetDataSource for the "People" smart marker. | Load a template workbook that contains smart markers, assign the LINQ‑filtered list as its data source, call Designer.Process, and save the populated file. | Verify that only rows matching the LINQ filter appear in the generated Excel sheet after processing the smart markers.
// Common Searches: how to filter data with LINQ before assigning to Aspose.Cells smart markers in C# | c# example of using WorkbookDesigner.SetDataSource with a LINQ‑filtered collection | smart markers rows missing after applying LINQ filter Aspose.Cells | bind filtered List<T> to smart marker named People using Aspose.Cells
// Tags: using LINQ with Aspose.Cells smart markers | assign filtered list to WorkbookDesigner data source | generate rows from filtered collection in smart markers | filtering data before smart marker processing Aspose.Cells | C# smart marker population with LINQ query

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // Sample data model
    // The example loads a template workbook containing smart markers, applies a LINQ filter to keep only persons older than 30, assigns the filtered List<Person> to the "People" smart marker via WorkbookDesigner.SetDataSource, processes the markers, and saves the result as Result.xlsx.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the workbook that contains smart markers (e.g., &amp;=People.Name, &amp;=People.Age)
            Workbook workbook = new Workbook("Template.xlsx");   // load-workbook rule

            // Create a WorkbookDesigner for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Original data collection
            List<Person> people = GetPeople();

            // LINQ filter: only include persons older than 30
            List<Person> filteredPeople = people
                .Where(p => p.Age > 30)
                .ToList();

            // Assign the filtered collection to the smart marker named "People"
            designer.SetDataSource("People", filteredPeople);   // chart-values / data-source rule

            // Process the smart markers – rows will be generated for each item in filteredPeople
            designer.Process();

            // Save the populated workbook
            workbook.Save("Result.xlsx");   // save-workbook rule
        }

        // Mock method to obtain sample data
        static List<Person> GetPeople()
        {
            return new List<Person>
            {
                new Person { Name = "Alice", Age = 28, City = "New York" },
                new Person { Name = "Bob",   Age = 35, City = "Chicago" },
                new Person { Name = "Carol", Age = 42, City = "Seattle" },
                new Person { Name = "Dave",  Age = 31, City = "Boston" }
            };
        }
    }
}
