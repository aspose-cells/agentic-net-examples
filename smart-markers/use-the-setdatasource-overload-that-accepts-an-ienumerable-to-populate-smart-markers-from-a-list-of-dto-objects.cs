// Title: Aspose.Cells: Populate Smart Markers from an IEnumerable&lt;T&gt; DTO List using SetDataSource
// Description: Demonstrates how to create a workbook template, add smart‑marker placeholders, bind a List&lt;Person&gt; (IEnumerable) to the marker name "Person" with WorkbookDesigner.SetDataSource, process the markers, and save the populated Excel file.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource IEnumerable | smart markers | C# Excel export | DTO list binding | populate template from List<T> | Excel automation | data‑driven workbook | IEnumerable overload
// Common Searches: bind List<T> to Aspose.Cells smart markers | WorkbookDesigner SetDataSource IEnumerable example | populate Excel from DTO collection using smart markers | Aspose.Cells smart marker list binding C# | export data to Excel with IEnumerable overload
// Developer Intent: Bind a collection of custom DTO objects to smart‑marker placeholders and generate a fully populated Excel workbook.
// Use Cases: Create an employee directory by feeding a List<Person> into a smart‑marker template. | Generate a sales ledger where each row reflects an Order DTO from a query result. | Export filtered database records to Excel by passing the result set as an IEnumerable to WorkbookDesigner.
// AI Prompts: Show a C# example of WorkbookDesigner.SetDataSource with an IEnumerable of custom objects. | Explain step‑by‑step how to bind a List<Person> to smart markers and save the workbook. | How can I customize smart‑marker syntax for nested DTO properties when using the IEnumerable overload?

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple DTO class representing a person
    // Demonstrates how to create a workbook template, add smart‑marker placeholders, bind a List&lt;Person&gt; (IEnumerable) to the marker name "Person" with WorkbookDesigner.SetDataSource, process the markers, and save the populated Excel file.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }

        public Person(string name, int age, string city)
        {
            Name = name;
            Age = age;
            City = city;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook that will serve as the template
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define column headers
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("City");

            // Insert smart markers that reference the "Person" data source
            sheet.Cells["A2"].PutValue("&=$Person.Name");
            sheet.Cells["B2"].PutValue("&=$Person.Age");
            sheet.Cells["C2"].PutValue("&=$Person.City");

            // Prepare a list of DTO objects (IEnumerable) to bind to the smart markers
            List<Person> persons = new List<Person>
            {
                new Person("John Doe", 30, "New York"),
                new Person("Jane Smith", 28, "London"),
                new Person("Sam Brown", 35, "Sydney")
            };

            // Initialize the WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Bind the list to the smart marker name "Person" using the IEnumerable overload
            designer.SetDataSource("Person", persons);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the populated workbook
            workbook.Save("SmartMarkersFromIEnumerable.xlsx");
        }
    }
}
