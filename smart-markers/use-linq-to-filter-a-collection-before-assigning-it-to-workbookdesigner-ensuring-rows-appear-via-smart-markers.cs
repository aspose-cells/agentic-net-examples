// Title: Filter a collection with LINQ before binding to WorkbookDesigner for smart‑marker rows (Aspose.Cells for .NET)
// Description: This C# example shows how to use LINQ to filter a List<Person> (e.g., Age > 30), bind the filtered list to a WorkbookDesigner smart‑marker named "Person", and generate a dynamic Excel sheet where each filtered record creates a new row. The workbook is saved as FilteredPersons.xlsx.
// Keywords: Aspose.Cells | C# | .NET | WorkbookDesigner | smart markers | LINQ filter | Excel export | SetDataSource | dynamic rows | code example
// Common Searches: Aspose.Cells LINQ filter before WorkbookDesigner | smart markers with filtered collection C# | how to bind LINQ result to WorkbookDesigner | generate Excel rows from LINQ query Aspose.Cells | filter list before processing smart markers
// Developer Intent: I need to apply a LINQ filter to a data collection and then bind only the filtered items to a WorkbookDesigner so that smart markers create rows solely for those items.
// Use Cases: Create an employee report that includes only staff members meeting age or department criteria. | Export a customized Excel sheet where business rules (e.g., sales > threshold) determine which records appear. | Generate department‑specific summaries by pre‑filtering data before smart‑marker processing.
// AI Prompts: Write C# code that filters a List<Person> with LINQ and binds the result to a WorkbookDesigner smart‑marker in Aspose.Cells. | Show how to add conditional formatting to the smart‑marker row after applying a LINQ filter. | Provide troubleshooting steps when filtered rows are missing after calling WorkbookDesigner.Process().

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerLinqDemo
{
    // Simple data class used as a data source
    // This C# example shows how to use LINQ to filter a List<Person> (e.g., Age > 30), bind the filtered list to a WorkbookDesigner smart‑marker named "Person", and generate a dynamic Excel sheet where each filtered record creates a new row. The workbook is saved as FilteredPersons.xlsx.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }

        public Person(string name, int age, string department)
        {
            Name = name;
            Age = age;
            Department = department;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare a collection of Person objects
            List<Person> allPersons = new List<Person>
            {
                new Person("John Doe", 28, "Sales"),
                new Person("Jane Smith", 35, "Marketing"),
                new Person("Bob Johnson", 42, "IT"),
                new Person("Alice Brown", 31, "HR")
            };

            // 2. Use LINQ to filter the collection (e.g., only persons older than 30)
            List<Person> filteredPersons = allPersons
                .Where(p => p.Age > 30)
                .ToList();

            // 3. Create a new workbook and set up smart markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("Department");

            // Smart marker row – the designer will repeat this row for each item in the data source
            sheet.Cells["A2"].PutValue("&Person.Name");
            sheet.Cells["B2"].PutValue("&Person.Age");
            sheet.Cells["C2"].PutValue("&Person.Department");

            // 4. Initialize WorkbookDesigner and bind the filtered collection
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the filtered list to the smart marker name "Person"
            designer.SetDataSource("Person", filteredPersons);

            // 5. Process the smart markers – rows will be generated for each filtered item
            designer.Process();

            // 6. Save the result
            workbook.Save("FilteredPersons.xlsx");
        }
    }
}
