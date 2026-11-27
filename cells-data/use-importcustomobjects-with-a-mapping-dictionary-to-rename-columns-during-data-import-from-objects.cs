// Title: Rename Excel column headers while importing custom objects with ImportCustomObjects in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells import custom objects with custom column names | How to map object property names to Excel headers using ImportCustomObjects | Replace default headers after ImportCustomObjects in Aspose.Cells | Use a dictionary to rename columns during Aspose.Cells data import | ImportCustomObjects with ImportTableOptions and column mapping in .NET
// Tags: ImportCustomObjects column header mapping | Aspose.Cells rename imported columns | C# import custom objects to Excel with custom headers | ImportTableOptions show field names | dictionary based column renaming Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ImportCustomObjectsDemo
{
    // Sample data class
    // The example creates a workbook, imports an array of Person objects using ImportCustomObjects, then replaces the automatically generated header cells with custom names defined in a dictionary, and finally saves the worksheet as an Excel file.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }

        public Person(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName  = lastName;
            Age       = age;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Prepare a collection of custom objects
                Person[] people = new Person[]
                {
                    new Person("John",   "Doe",      30),
                    new Person("Alice",  "Smith",    25),
                    new Person("Robert", "Johnson",  40)
                };

                // 3. Define a column‑mapping dictionary to rename the properties
                //    Key   = original property name
                //    Value = desired column header in the worksheet
                Dictionary<string, string> columnMapping = new Dictionary<string, string>()
                {
                    { "FirstName", "First Name" },
                    { "LastName",  "Surname"    },
                    { "Age",       "Years"      }
                };

                // 4. Set import options (show column headers)
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true   // Show the column headers (original names)
                };

                // 5. Import the custom objects starting at cell A1 (row 0, column 0)
                //    The method returns the number of rows imported (including header row)
                int rowsImported = sheet.Cells.ImportCustomObjects(
                    people,
                    0,               // startRow
                    0,               // startColumn
                    importOptions    // import options
                );

                // 6. Rename the header cells according to the mapping dictionary
                //    The header row is the first imported row (row index 0)
                string[] propertyOrder = { "FirstName", "LastName", "Age" };
                for (int col = 0; col < propertyOrder.Length; col++)
                {
                    string originalName = propertyOrder[col];
                    if (columnMapping.TryGetValue(originalName, out string mappedName))
                    {
                        sheet.Cells[0, col].PutValue(mappedName);
                    }
                }

                Console.WriteLine($"Successfully imported {rowsImported - 1} data rows (plus header).");

                // 7. Save the workbook
                string outputPath = "ImportCustomObjectsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the import process:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
