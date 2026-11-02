// Title: Set CheckMergedCells = true before using ImportCustomObjects to preserve merged ranges in Aspose.Cells (C#)
// Common Searches: how to preserve merged cells when importing a list of objects with Aspose.Cells C# | Aspose.Cells ImportCustomObjects merged cell ranges not populated | enable CheckMergedCells option before ImportCustomObjects in C# | C# example for setting CheckMergedCells true with Aspose.Cells | ImportCustomObjects loses merged cells Aspose.Cells workaround
// Tags: ImportCustomObjects merged cell handling Aspose.Cells | CheckMergedCells property C# Aspose.Cells | populate merged ranges during object import | Aspose.Cells worksheet import custom objects | C# Excel merged cells import option

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example demonstrates creating a workbook, defining a list of custom Person objects, and importing them with ImportCustomObjects. It highlights the need to set the CheckMergedCells property to true before the import so that any pre‑existing merged cell ranges are correctly populated in the resulting Excel file.
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

            // Sample custom objects to import
            List<Person> people = new List<Person>
            {
                new Person { Name = "John", Age = 30 },
                new Person { Name = "Jane", Age = 25 }
            };

            // Set import options to include field names as headers
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true
            };

            // Import the custom objects starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportCustomObjects(people.ToArray(), 0, 0, importOptions);

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Definition of the custom object
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
