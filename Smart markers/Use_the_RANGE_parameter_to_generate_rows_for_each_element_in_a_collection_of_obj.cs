using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeExample
{
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            Workbook workbook = new Workbook("template.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 25 },
                new Person { Name = "Carol", Age = 28 }
            };

            int startRow = 0;
            int startColumn = 0;
            int totalRows = people.Count + 1;
            int totalColumns = 2;

            AsposeRange dataRange = cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            dataRange[0, 0].PutValue("Name");
            dataRange[0, 1].PutValue("Age");

            for (int i = 0; i < people.Count; i++)
            {
                dataRange[i + 1, 0].PutValue(people[i].Name);
                dataRange[i + 1, 1].PutValue(people[i].Age);
            }

            workbook.Save("output.xlsx");
        }
    }
}