// Title: Import a List of custom objects into an Excel worksheet and format the Date column as dd-MMM-yyyy with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells import list of objects into worksheet with headers | How to apply custom date format dd-MMM-yyyy to a column using Aspose.Cells | Set cell style for DateTime column after importing custom objects in Aspose.Cells | Aspose.Cells format BirthDate column while importing Person objects in C#
// Tags: import custom objects Aspose.Cells C# | custom number format dd-MMM-yyyy Aspose.Cells | apply cell style to DateTime column Excel C# | set worksheet column style Aspose.Cells | format BirthDate column Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Demonstrates creating a workbook, importing a List<Person> as a table with headers, applying a custom number format "dd-MMM-yyyy" to the BirthDate column, and saving the file as an XLSX using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Sample data
            var data = new List<Person>()
            {
                new Person { Name = "Alice", BirthDate = new DateTime(1990, 5, 12) },
                new Person { Name = "Bob", BirthDate = new DateTime(1985, 11, 23) },
                new Person { Name = "Charlie", BirthDate = new DateTime(2000, 1, 5) }
            };

            // Import custom objects with headers (first row)
            var importOptions = new ImportTableOptions
            {
                // Show property names as column headers
                IsFieldNameShown = true
            };
            sheet.Cells.ImportCustomObjects(data.ToArray(), 0, 0, importOptions);

            // Apply custom number format to the BirthDate column (B)
            int birthDateColumnIndex = 1; // zero‑based index for column B
            var style = workbook.CreateStyle();
            style.Custom = "dd-MMM-yyyy";

            // Apply the style to each cell in the BirthDate column (skip header row)
            int lastRow = sheet.Cells.MaxDataRow;
            for (int row = 1; row <= lastRow; row++)
            {
                var cell = sheet.Cells[row, birthDateColumnIndex];
                cell.SetStyle(style);
            }

            // Save the workbook
            workbook.Save("CustomObjectsWithFormattedDates.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Custom class definition
    public class Person
    {
        public string Name { get; set; } = null!;
        public DateTime BirthDate { get; set; }
    }
}
