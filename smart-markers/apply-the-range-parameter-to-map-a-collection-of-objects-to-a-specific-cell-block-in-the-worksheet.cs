// Title: Import a C# object list into a specific Excel range using Aspose.Cells ImportCustomObjects
// Common Searches: how to import a list of C# objects into a defined cell range with Aspose.Cells | Aspose.Cells ImportCustomObjects example for mapping objects to B2 cell block | C# import collection into Excel worksheet without field name headers using Aspose.Cells | specify start row and column when importing custom objects with Aspose.Cells | populate Excel range from a List<Person> using Aspose.Cells in .NET
// Tags: ImportCustomObjects start row column Aspose.Cells | C# map object collection to Excel range | Aspose.Cells hide field names import | populate Excel worksheet from list C# | custom object import to .xlsx using Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example creates a new Workbook, selects the first worksheet, and sets B2 (row index 1, column index 1) as the starting cell. It then imports an array of Person objects into the worksheet with ImportCustomObjects, using ImportTableOptions to suppress automatic field name headers. Manual headers are added, and the workbook is saved as MappedData.xlsx.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public double Salary { get; set; }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Sample collection of objects to map
            var data = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30, Salary = 50000 },
                new Person { Name = "Bob",     Age = 25, Salary = 45000 },
                new Person { Name = "Charlie", Age = 35, Salary = 60000 }
            };

            // Define the start cell for the range (e.g., B2)
            int startRow = 1;      // zero‑based index (row 2)
            int startColumn = 1;   // zero‑based index (column B)

            // Set import options (do not import field names automatically)
            var importOptions = new ImportTableOptions
            {
                IsFieldNameShown = false
            };

            // Import the collection into the worksheet using the Range parameters.
            // This maps each object's properties to consecutive columns.
            sheet.Cells.ImportCustomObjects(data.ToArray(), startRow, startColumn, importOptions);

            // Optionally add headers for clarity
            sheet.Cells[startRow, startColumn - 1].PutValue("Name");
            sheet.Cells[startRow, startColumn].PutValue("Age");
            sheet.Cells[startRow, startColumn + 1].PutValue("Salary");

            // Save the workbook to a file
            string outputPath = "MappedData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
