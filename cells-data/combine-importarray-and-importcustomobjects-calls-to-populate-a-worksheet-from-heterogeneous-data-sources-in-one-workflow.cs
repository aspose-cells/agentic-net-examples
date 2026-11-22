// Title: Import a primitive array and a list of custom objects into the same worksheet using Aspose.Cells for C#
// Common Searches: how to import a two dimensional object array into an Aspose.Cells worksheet in C# | append a list of custom objects after importing array data with Aspose.Cells | combine ImportArray and ImportCustomObjects calls in one Aspose.Cells workbook | populate Excel sheet with mixed primitive array and POCO list using Aspose.Cells C# | set start row for ImportCustomObjects after ImportArray in Aspose.Cells
// Tags: ImportArray primitive data Aspose.Cells C# | ImportCustomObjects POCO list Aspose.Cells | heterogeneous data worksheet population Aspose.Cells | combine array and object import Aspose.Cells | start row for ImportCustomObjects Aspose.Cells | save workbook as XLSX Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, imports a two‑dimensional primitive array row by row with Cells.ImportArray, then adds a list of Person objects using Cells.ImportCustomObjects (including column headers), and saves the result as HeterogeneousData.xlsx.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample heterogeneous data: primitive array
            object[,] primitiveData = new object[,]
            {
                { "Product", "Quantity", "Price" },
                { "Apple", 10, 0.5 },
                { "Banana", 20, 0.3 }
            };

            // Import the primitive array row by row (ImportArray rule)
            int rows = primitiveData.GetLength(0);
            int cols = primitiveData.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                // Convert each cell to string for the ImportArray(string[] ...) overload
                string[] rowData = new string[cols];
                for (int j = 0; j < cols; j++)
                {
                    rowData[j] = primitiveData[i, j]?.ToString() ?? string.Empty;
                }

                // Import the current row at the appropriate position
                sheet.Cells.ImportArray(rowData, i, 0, false);
            }

            // Custom data source
            List<Person> people = new List<Person>
            {
                new Person { Id = 1, Name = "Alice", Age = 30 },
                new Person { Id = 2, Name = "Bob", Age = 25 }
            };

            // Determine start row after the array data
            int startRow = rows;

            // Import custom objects (ImportCustomObjects rule)
            var importOptions = new ImportTableOptions
            {
                // Show field names as column headers
                IsFieldNameShown = true
            };
            sheet.Cells.ImportCustomObjects(people, startRow, 0, importOptions);

            // Save the workbook (save rule)
            string outputPath = "HeterogeneousData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Custom object definition
    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty; // non‑null default
        public int Age { get; set; }
    }
}
