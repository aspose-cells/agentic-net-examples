using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportExample
{
    // Define a simple data class
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Gender { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Prepare sample data
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30, Gender = "Female" },
                new Person { Name = "Bob",     Age = 25, Gender = "Male"   },
                new Person { Name = "Charlie", Age = 35, Gender = "Male"   }
            };

            // Get the factory for building ICellsDataTable from custom objects (factory rule)
            CellsDataTableFactory factory = workbook.CellsDataTableFactory;

            // Create an ICellsDataTable instance from the list (no custom code needed)
            ICellsDataTable dataTable = factory.GetInstance(people, true);

            // Configure import options:
            // - Show field names (column headers)
            // - Map only the Name and Age columns (skip Gender) by specifying column indexes
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                ColumnIndexes = new int[] { 0, 1 } // 0 = Name, 1 = Age
            };

            // Import the data into the first worksheet starting at cell A1 (row 0, column 0)
            // (import rule)
            workbook.Worksheets[0].Cells.ImportData(dataTable, 0, 0, importOptions);

            // Save the workbook to an XLSX file (save rule)
            workbook.Save("PeopleData.xlsx");
        }
    }
}