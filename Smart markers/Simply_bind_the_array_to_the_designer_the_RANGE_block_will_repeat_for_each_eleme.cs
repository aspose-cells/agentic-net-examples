using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsArrayBindingDemo
{
    // Sample data class that will be bound to the smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook (must contain a range block with smart markers like &Person.Name, &Person.Age)
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the designer with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Create an array (List) of Person objects to bind
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30 },
                new Person { Name = "Bob",     Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            };

            // Bind the array to the designer using the variable name defined in the smart markers ("Person")
            designer.SetDataSource("Person", persons);

            // Process the smart markers; the range block will repeat for each element in the array
            designer.Process();

            // Save the populated workbook
            workbook.Save("output.xlsx");
        }
    }
}