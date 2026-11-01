// Title: How to rename Excel column headers when importing a list of custom objects with ImportCustomObjects in Aspose.Cells for C#
// Common Searches: Aspose.Cells C# import list of objects and set custom column headers | rename generated headers after ImportCustomObjects using a dictionary in C# | map class property names to friendly Excel column titles with Aspose.Cells | C# example of ImportCustomObjects with custom header mapping | change header row after importing custom objects in an Aspose.Cells workbook
// Tags: ImportCustomObjects custom header mapping | C# Aspose.Cells rename column headers | dictionary based Excel header customization Aspose.Cells | field name shown option ImportTableOptions | save imported data to Excel workbook C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportExample
{
    // Define a custom class whose objects will be imported.
    // The example creates a workbook, imports a List<Person> into the first worksheet using ImportCustomObjects with headers enabled, then replaces the automatically generated column headers with custom titles defined in a dictionary, and finally saves the file as ImportedData.xlsx.
    public class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName  { get; set; } = string.Empty;
        public int    Age       { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Get the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare a collection of custom objects to import.
                List<Person> people = new List<Person>
                {
                    new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
                    new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 },
                    new Person { FirstName = "Alice", LastName = "Brown", Age = 28 }
                };

                // Set import options to include headers (property names).
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    // Show property names as column headers.
                    IsFieldNameShown = true
                };

                // Import the custom objects starting at cell A1 (row 0, column 0).
                sheet.Cells.ImportCustomObjects(people, 0, 0, importOptions);

                // Rename the automatically generated headers using the desired mapping.
                var columnMapping = new Dictionary<string, string>
                {
                    { "FirstName", "First Name" },
                    { "LastName",  "Last Name"  },
                    { "Age",       "Years"      }
                };

                // The header row is the first row (row index 0).
                // Iterate through the properties to replace header text.
                int colIndex = 0;
                foreach (var prop in typeof(Person).GetProperties())
                {
                    if (columnMapping.TryGetValue(prop.Name, out string newHeader))
                    {
                        sheet.Cells[0, colIndex].PutValue(newHeader);
                    }
                    colIndex++;
                }

                // Save the workbook.
                workbook.Save("ImportedData.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
