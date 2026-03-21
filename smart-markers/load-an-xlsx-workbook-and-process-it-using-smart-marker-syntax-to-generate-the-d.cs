using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    // Simple data class used as a data source for smart markers
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
            // Load the template workbook that contains smart markers (e.g., &Person.Name, &Person.Age, &Person.Department)
            Workbook templateWorkbook = new Workbook("template.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Prepare sample data source
            List<Person> employees = new List<Person>
            {
                new Person("John Doe", 30, "Sales"),
                new Person("Jane Smith", 28, "Marketing"),
                new Person("Bob Johnson", 35, "Engineering")
            };

            // Bind the data source to the smart marker name "Person"
            designer.SetDataSource("Person", employees);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the processed workbook to the desired output file
            designer.Workbook.Save("output.xlsx");
        }
    }
}