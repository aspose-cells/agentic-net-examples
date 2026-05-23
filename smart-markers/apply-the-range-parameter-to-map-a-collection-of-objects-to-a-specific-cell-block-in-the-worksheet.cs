using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsRangeMappingDemo
{
    // Sample data class
    public class Person
    {
        public string Name { get; set; } = string.Empty;   // initialize to avoid nullable warning
        public int Age { get; set; }
    }

    public class Program
    {
        // Alias to avoid conflict with System.Range
        private static readonly Type AsposeRangeType = typeof(Aspose.Cells.Range);

        public static void Main()
        {
            try
            {
                // Create a collection of objects to map
                List<Person> people = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                };

                // Initialize a new workbook and get the first worksheet's cells
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the size of the range (header row + data rows, 2 columns)
                int totalRows = people.Count + 1; // +1 for header
                int totalColumns = 2; // Name and Age

                // Create a 2‑dimensional array to hold the values
                object[,] values = new object[totalRows, totalColumns];

                // Set header values
                values[0, 0] = "Name";
                values[0, 1] = "Age";

                // Populate the array with data from the collection
                for (int i = 0; i < people.Count; i++)
                {
                    values[i + 1, 0] = people[i].Name;
                    values[i + 1, 1] = people[i].Age;
                }

                // Create a range that starts at cell A1 and spans the required rows and columns
                // Use fully qualified type to avoid ambiguity with System.Range
                Aspose.Cells.Range range = cells.CreateRange(0, 0, totalRows, totalColumns);

                // Assign the 2‑D array to the range's Value property
                range.Value = values;

                // Add the range to the worksheet's Cells collection (optional, demonstrates AddRange usage)
                cells.AddRange(range);

                // Save the workbook to a file
                string outputPath = "PeopleRangeMapping.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}