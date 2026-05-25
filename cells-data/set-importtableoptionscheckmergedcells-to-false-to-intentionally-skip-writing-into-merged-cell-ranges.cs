using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a merged cell range D4:D5 (row 3, column 3)
        cells.Merge(3, 3, 2, 1);
        cells[3, 3].PutValue("MergedValue");

        // Sample data to import
        List<Customer> customers = new List<Customer>
        {
            new Customer { CustomerId = 1, Name = "Alice" },
            new Customer { CustomerId = 2, Name = "Bob" },
            new Customer { CustomerId = 3, Name = "Charlie" }
        };

        // Configure import options with CheckMergedCells set to false
        ImportTableOptions options = new ImportTableOptions
        {
            IsFieldNameShown = false,
            InsertRows = true,
            CheckMergedCells = false // Intentionally skip checking merged cells
        };

        // Import the custom objects starting at the merged cell location
        // Because CheckMergedCells is false, the import will write into the merged range
        cells.ImportCustomObjects(customers, 3, 3, options);

        // Save the workbook
        workbook.Save("MergedSkipDemo.xlsx", SaveFormat.Xlsx);
    }

    // Simple POCO class used for import
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; }
    }
}