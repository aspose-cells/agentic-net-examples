// Title: Import a collection of custom objects into an Excel worksheet with column headers starting at row 2 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# import list of objects into Excel starting at row 2 with headers | how to map class properties to Excel columns using ImportCustomObjects in Aspose.Cells | C# example for importing custom objects array into worksheet with field names shown | save a collection of Person objects to an .xlsx file using Aspose.Cells ImportTableOptions | import custom objects to first worksheet at row index 1 column index 0 Aspose.Cells
// Tags: import custom objects to Excel Aspose.Cells | map class properties to worksheet columns C# | ImportCustomObjects with field name headers | start import at specific row Aspose.Cells | ImportTableOptions show field names | save workbook to .xlsx using Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportExample
{
    // Custom object whose properties will be mapped to worksheet columns
    // The example creates a new workbook, defines a Person class, builds a list of Person objects, configures ImportTableOptions to display field names as column headers, and uses ImportCustomObjects to write the array into the first worksheet beginning at row 2, column 1, then saves the file as PeopleReport.xlsx.
    public class Person
    {
        public string Name { get; set; } = string.Empty;   // Column A
        public int Age { get; set; }                       // Column B
        public DateTime BirthDate { get; set; }            // Column C
        public double Salary { get; set; }                 // Column D
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare a collection of custom objects
                List<Person> people = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30, BirthDate = new DateTime(1993, 5, 12), Salary = 75000 },
                    new Person { Name = "Bob",   Age = 45, BirthDate = new DateTime(1978, 11, 3), Salary = 92000 },
                    new Person { Name = "Carol", Age = 28, BirthDate = new DateTime(1995, 2, 20), Salary = 68000 }
                };

                // Set import options to include column headers
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true
                };

                // Import the collection starting at row 2 (index 1) and column 1 (index 0)
                sheet.Cells.ImportCustomObjects(people.ToArray(), 1, 0, importOptions);

                // Save the workbook to a file
                string outputPath = "PeopleReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
