using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerObjectPopulation
{
    // Sample data classes used as data sources for smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Wife Wife { get; set; }

        public Person(string name, int age, Wife wife)
        {
            Name = name;
            Age = age;
            Wife = wife;
        }
    }

    public class Wife
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Wife(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers (e.g., &=$Person.Name, &=$Person.Wife.Name)
            Workbook templateWorkbook = new Workbook("template.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Prepare a collection of Person objects as the data source
            List<Person> persons = new List<Person>
            {
                new Person("John Doe", 35, new Wife("Jane Doe", 33)),
                new Person("Alice Smith", 28, new Wife("Bob Smith", 30))
            };

            // Bind the collection to the smart marker variable "Person"
            designer.SetDataSource("Person", persons);

            // Process the smart markers and populate the worksheet with the object data
            designer.Process();

            // Save the populated workbook
            designer.Workbook.Save("output.xlsx");
        }
    }
}