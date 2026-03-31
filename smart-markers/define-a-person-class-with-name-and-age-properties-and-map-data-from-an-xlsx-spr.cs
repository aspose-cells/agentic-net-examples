using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsMappingDemo
{
    // Simple POCO representing a person
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source Excel file (must exist)
            string inputPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range (including header row)
            int firstRow = 0;                     // header row (A1)
            int firstColumn = 0;                  // first column (A)
            int totalRows = cells.MaxDataRow + 1; // MaxDataRow is zero‑based
            int totalColumns = cells.MaxDataColumn + 1;

            // Export the range to a DataTable (header row will be used as column names)
            DataTable dt = cells.ExportDataTable(firstRow, firstColumn, totalRows, totalColumns, true);

            // Determine column indices (fallback to positional indices if names are missing)
            int nameColIndex = dt.Columns.Contains("Name") ? dt.Columns.IndexOf("Name") : 0;
            int ageColIndex = dt.Columns.Contains("Age") ? dt.Columns.IndexOf("Age") : 1;

            // Map each DataRow to a Person instance
            List<Person> people = new List<Person>();
            foreach (DataRow row in dt.Rows)
            {
                string name = row[nameColIndex]?.ToString() ?? string.Empty;
                int age = 0;
                int.TryParse(row[ageColIndex]?.ToString(), out age);

                people.Add(new Person { Name = name, Age = age });
            }

            // Example usage: output the imported persons to console
            Console.WriteLine("Imported persons:");
            foreach (var p in people)
            {
                Console.WriteLine($"Name: {p.Name}, Age: {p.Age}");
            }

            // Save the workbook after processing
            workbook.Save("output.xlsx");
        }
    }
}