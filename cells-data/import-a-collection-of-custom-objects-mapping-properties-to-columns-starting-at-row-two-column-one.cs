// Title: Import a list of custom C# objects into an Excel worksheet at row 2, column A with column headers using Aspose.Cells
// Common Searches: Aspose.Cells C# import custom object collection starting at row 2 column A | How to map object properties to Excel columns with ImportCustomObjects in .NET | Include field names as headers when importing a List<T> into Excel using Aspose.Cells | Save a workbook after importing custom objects with Aspose.Cells for .NET | Set ImportTableOptions.IsFieldNameShown true in Aspose.Cells import
// Tags: ImportCustomObjects with start row and column Aspose.Cells | C# collection to Excel mapping object properties | ImportTableOptions field name header Aspose.Cells | Export List<Person> to XLSX using Aspose.Cells | Specify import start cell in Aspose.Cells worksheet

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportExample
{
    // Define a custom object with properties to map to columns
    // The example creates a new Workbook, defines a Person class, builds a List<Person>, configures ImportTableOptions to display property names as column headers, and calls sheet.Cells.ImportCustomObjects to write the collection beginning at row 2 (index 1) and column A (index 0). The workbook is then saved as CustomObjectsImport.xlsx.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Email { get; set; } = string.Empty;
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
                    new Person { Name = "Alice", Age = 30, Email = "alice@example.com" },
                    new Person { Name = "Bob",   Age = 25, Email = "bob@example.com" },
                    new Person { Name = "Carol", Age = 28, Email = "carol@example.com" }
                };

                // Set import options to include column headers (property names)
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true
                };

                // Import the collection starting at row 2 (index 1) and column 1 (index 0)
                sheet.Cells.ImportCustomObjects(people.ToArray(), 1, 0, importOptions);

                // Save the workbook to a file
                workbook.Save("CustomObjectsImport.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
