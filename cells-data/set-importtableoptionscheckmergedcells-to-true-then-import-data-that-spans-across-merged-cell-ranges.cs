// Title: Import Custom Objects with Merged Cells Preserved – Aspose.Cells C# Example
// Description: Demonstrates how to set ImportTableOptions.CheckMergedCells to true, merge cells B1:B2, add a header, and import a List<Customer> into an Aspose.Cells worksheet without breaking the merged layout, then save the workbook as an XLSX file.
// Keywords: Aspose.Cells | ImportCustomObjects | ImportTableOptions | CheckMergedCells | merged cells | C# | .NET | Excel worksheet | preserve merged cells | data import | template population | List<Customer>
// Common Searches: Aspose.Cells preserve merged cells on import | ImportCustomObjects with merged header C# | CheckMergedCells true example | How to keep merged cells when importing data with Aspose.Cells | Import list of objects into Excel template without breaking merged cells
// Developer Intent: Enable CheckMergedCells and import a collection of objects so existing merged ranges stay intact.
// Use Cases: Populate a report template that uses merged cells for section titles. | Add rows of customer data to a worksheet that already has a merged header row. | Automate data entry into Excel forms where layout relies on merged cells.
// AI Prompts: Write C# code that imports a List<T> into an Aspose.Cells worksheet while preserving merged cells using ImportTableOptions.CheckMergedCells. | Show how to configure ImportTableOptions to ignore merged cells during an import with Aspose.Cells. | Provide an example of importing data into a worksheet with multiple merged regions and custom column mapping in Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to set ImportTableOptions.CheckMergedCells to true, merge cells B1:B2, add a header, and import a List<Customer> into an Aspose.Cells worksheet without breaking the merged layout, then save the workbook as an XLSX file.
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

        // Merge cells B1:B2 (row 0, column 1, spanning 2 rows, 1 column)
        cells.Merge(0, 1, 2, 1);
        // Set a value in the merged cell to verify it remains after import
        cells[0, 1].PutValue("MergedHeader");

        // Prepare sample data to import
        List<Customer> customers = new List<Customer>
        {
            new Customer { CustomerId = 1, Name = "Alice" },
            new Customer { CustomerId = 2, Name = "Bob" },
            new Customer { CustomerId = 3, Name = "Charlie" }
        };

        // Configure import options with CheckMergedCells enabled
        ImportTableOptions options = new ImportTableOptions
        {
            IsFieldNameShown = true,   // import column headers
            InsertRows = true,         // insert rows instead of overwriting
            CheckMergedCells = true    // preserve merged cells during import
        };

        // Import the custom objects starting at cell A1 (row 0, column 0)
        cells.ImportCustomObjects(customers, 0, 0, options);

        // Save the workbook to a file
        workbook.Save("MergedImportDemo.xlsx", SaveFormat.Xlsx);
    }
}
