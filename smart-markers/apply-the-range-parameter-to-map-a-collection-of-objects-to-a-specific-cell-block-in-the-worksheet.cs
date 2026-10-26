// Title: Import a C# object collection into a specific worksheet cell block (B2) using Aspose.Cells ImportCustomObjects
// Common Searches: how to use ImportCustomObjects to start importing at cell B2 in Aspose.Cells C# | C# code to map a List<Person> to a defined range in an Excel sheet with Aspose.Cells | preserve column headers while importing objects into a worksheet range using Aspose.Cells | set start row and column parameters for ImportCustomObjects in Aspose.Cells .NET
// Tags: ImportCustomObjects start row column Aspose.Cells | map object collection to worksheet range C# | preserve formatting during object import Aspose.Cells | smart markers custom object import Excel | export List<Person> to Excel using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, defines a List<Person> with Name, Age, and City, sets the starting cell to B2 (row 1, column 1), and calls sheet.Cells.ImportCustomObjects to import the collection into the worksheet with column headers and formatting preserved, then saves the file as MappedObjects.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Sample collection of objects to map
                var people = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30, City = "New York" },
                    new Person { Name = "Bob",   Age = 25, City = "London"   },
                    new Person { Name = "Carol", Age = 28, City = "Paris"    }
                };

                // Define the fields (properties) to import, in the order they should appear in the worksheet
                string[] fields = { "Name", "Age", "City" };

                // Starting cell for the block (B2 -> row index 1, column index 1)
                int startRow = 1;
                int startColumn = 1;

                // Import the collection into the worksheet starting at the defined cell
                // Parameters: collection, fields, includeColumnNames, startRow, startColumn,
                // insertRows, preserveFormatting, dateTimeFormat, convertDateTime
                sheet.Cells.ImportCustomObjects(
                    people,
                    fields,
                    true,          // include column names
                    startRow,
                    startColumn,
                    0,             // insertRows (0 = no extra rows)
                    true,          // preserve formatting
                    null,          // dateTimeFormat (null = default)
                    false);        // convertDateTime

                // Save the workbook to a file
                string outputPath = "MappedObjects.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Helper class representing the objects to be mapped
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string City { get; set; } = string.Empty;
    }
}
