using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMergedImportDemo
{
    // Simple POCO representing a customer
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a merged region D4:D5 (row index 3, column index 3, 2 rows, 1 column)
            worksheet.Cells.Merge(3, 3, 2, 1);
            worksheet.Cells[3, 3].PutValue("MergedValue");

            // Prepare sample custom objects
            List<Customer> customers = new List<Customer>
            {
                new Customer { CustomerId = 1, Name = "Alice" },
                new Customer { CustomerId = 2, Name = "Bob" },
                new Customer { CustomerId = 3, Name = "Charlie" }
            };

            // Configure import options with merged‑cell checking enabled
            ImportTableOptions options = new ImportTableOptions
            {
                IsFieldNameShown = false,
                InsertRows = true,
                CheckMergedCells = true
            };

            // Import the custom objects starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportCustomObjects(customers, 0, 0, options);

            // Verify that the original merged region is still present
            Cell mergedCell = worksheet.Cells["D4"]; // D4 corresponds to row 3, column 3
            Console.WriteLine($"Merged cell value: {mergedCell.StringValue}");
            Console.WriteLine($"Is cell D4 merged? {mergedCell.IsMerged}");

            // Enumerate all merged areas in the worksheet
            CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();
            Console.WriteLine($"Total merged areas: {mergedAreas.Length}");
            foreach (CellArea area in mergedAreas)
            {
                Console.WriteLine($"Merged area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, EndRow={area.EndRow}, EndColumn={area.EndColumn}");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("MergedImportDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}