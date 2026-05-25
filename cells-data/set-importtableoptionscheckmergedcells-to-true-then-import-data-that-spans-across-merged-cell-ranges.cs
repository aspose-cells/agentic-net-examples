using System;
using System.Collections.Generic;
using Aspose.Cells;

public class Customer
{
    public int CustomerId { get; set; }
    public string Name { get; set; }
}

public class ImportMergedCellsDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells D4:D5 (zero‑based indices: row 3, column 3)
        cells.Merge(3, 3, 2, 1);
        cells[3, 3].PutValue("MergedHeader");

        // Sample data to import
        List<Customer> customers = new List<Customer>
        {
            new Customer { CustomerId = 1, Name = "Alice" },
            new Customer { CustomerId = 2, Name = "Bob" },
            new Customer { CustomerId = 3, Name = "Charlie" }
        };

        // Configure import options with CheckMergedCells enabled
        ImportTableOptions options = new ImportTableOptions
        {
            IsFieldNameShown = false,   // Do not import property names as header
            InsertRows = true,          // Insert rows instead of overwriting existing ones
            CheckMergedCells = true     // Preserve merged cells during import
        };

        // Import the custom objects starting at cell A1 (row 0, column 0)
        cells.ImportCustomObjects(customers, 0, 0, options);

        // Save the workbook
        workbook.Save("MergedImportDemo.xlsx", SaveFormat.Xlsx);
    }
}