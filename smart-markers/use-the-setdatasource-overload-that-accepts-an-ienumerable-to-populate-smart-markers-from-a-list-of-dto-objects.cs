using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // DTO class representing a person
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
            // 1. Create a list of DTO objects
            List<Person> persons = new List<Person>
            {
                new Person("John Doe", 30, "New York"),
                new Person("Jane Smith", 28, "London"),
                new Person("Sam Brown", 35, "Sydney")
            };

            // 2. Create a new workbook (or load a template if you have one)
            Workbook workbook = new Workbook();

            // 3. Insert smart markers into the worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("City");
            // Data row with smart markers (the variable name is "Person")
            sheet.Cells["A2"].PutValue("&=$Person.Name");
            sheet.Cells["B2"].PutValue("&=$Person.Age");
            sheet.Cells["C2"].PutValue("&=$Person.City");

            // 4. Initialize WorkbookDesigner and bind the IEnumerable data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Use the overload SetDataSource(string, object) where the object is an IEnumerable
            designer.SetDataSource("Person", persons);

            // 5. Process the smart markers
            designer.Process();

            // 6. Save the populated workbook
            workbook.Save("SmartMarkersFromDTO.xlsx");
        }
    }
}