// Title: Aspose.Cells .NET – Preserve Merged Cells When Importing Custom Objects (CheckMergedCells)
// Description: This example creates a workbook, merges cells D4:D5, builds a List<Customer>, configures ImportTableOptions with CheckMergedCells = true (and InsertRows enabled, headers hidden), imports the collection starting at A1, and saves the file as XLSX, ensuring the merged range remains intact after the import.
// Keywords: Aspose.Cells | ImportCustomObjects | CheckMergedCells | merged cells | custom objects import | C# | .NET | ImportTableOptions | worksheet data load | Excel automation | GitHub example | coding‑agent snippet
// Common Searches: Aspose.Cells keep merged cells after ImportCustomObjects | CheckMergedCells option C# example | Import custom objects into worksheet with merged headers | .NET preserve merged ranges during data import | How to use ImportTableOptions CheckMergedCells
// Developer Intent: Enable the CheckMergedCells flag before calling ImportCustomObjects so existing merged ranges are not broken during the import.
// Use Cases: Load a collection of business objects into a sheet that already contains a merged title row. | Append rows to a report while maintaining merged header cells. | Automate Excel generation where merged cells define sections that must stay unchanged after bulk data insertion.
// AI Prompts: Generate C# code that imports a list of objects into an Aspose.Cells worksheet without destroying merged cells. | Explain how the CheckMergedCells property influences merged ranges during ImportCustomObjects in Aspose.Cells for .NET. | Show a step‑by‑step tutorial for preserving merged cells when importing custom objects with ImportTableOptions.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCheckMergedCellsDemo
{
    // Sample custom object to be imported
    // This example creates a workbook, merges cells D4:D5, builds a List<Customer>, configures ImportTableOptions with CheckMergedCells = true (and InsertRows enabled, headers hidden), imports the collection starting at A1, and saves the file as XLSX, ensuring the merged range remains intact after the import.
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a merged cell range D4:D5 (zero‑based indices: row 3, column 3)
            worksheet.Cells.Merge(3, 3, 2, 1);
            worksheet.Cells[3, 3].PutValue("MergedValue");

            // Prepare a list of custom objects to import
            List<Customer> customers = new List<Customer>
            {
                new Customer { CustomerId = 1, Name = "Alice" },
                new Customer { CustomerId = 2, Name = "Bob" },
                new Customer { CustomerId = 3, Name = "Charlie" }
            };

            // Set import options and enable checking of merged cells
            ImportTableOptions options = new ImportTableOptions
            {
                CheckMergedCells = true,   // Important: preserve merged cell ranges
                InsertRows = true,         // Add rows if needed
                IsFieldNameShown = false   // Do not import property names as header
            };

            // Import the custom objects starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportCustomObjects((ICollection)customers, 0, 0, options);

            // Save the workbook to an XLSX file
            workbook.Save("CheckMergedCellsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
