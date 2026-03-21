using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsRangeExample
{
    // Simple data class
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (template) – lifecycle rule: load
            Workbook workbook = new Workbook("Template.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample collection of objects to be written into the sheet
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            };

            // Starting row where data will be inserted (0‑based index)
            int startRow = 1; // e.g., row 2 in Excel (assuming row 1 has headers)

            // Loop through the collection and create a one‑row range for each object
            for (int i = 0; i < people.Count; i++)
            {
                // Calculate the row index for the current object
                int currentRow = startRow + i;

                // Create a range that spans one row and two columns (Name, Age)
                // Using Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns)
                Aspose.Cells.Range rowRange = cells.CreateRange(currentRow, 0, 1, 2);

                // Populate the range with the object's data
                rowRange.Value = new object[,] { { people[i].Name, people[i].Age } };
            }

            // Save the modified workbook – lifecycle rule: save
            workbook.Save("Result.xlsx");
        }
    }
}