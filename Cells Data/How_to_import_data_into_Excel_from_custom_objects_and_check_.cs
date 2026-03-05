using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportCustomObjectsDemo
{
    // Sample custom class to be imported
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();                     // create
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 2. Create a merged area for demonstration (e.g., D4:D5)
            cells.Merge(3, 3, 2, 1);                               // merge rows 4-5, column D
            cells[3, 3].PutValue("MergedHeader");                  // put a value in the merged cell

            // 3. Prepare a list of custom objects to import
            List<Customer> customers = new List<Customer>
            {
                new Customer { CustomerId = 1, Name = "Alice" },
                new Customer { CustomerId = 2, Name = "Bob" },
                new Customer { CustomerId = 3, Name = "Charlie" }
            };

            // 4. Configure import options
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = false,   // do not import property names as header
                InsertRows = true,          // insert rows if needed
                CheckMergedCells = true     // enable checking of merged cells
            };

            // 5. Import the custom objects starting at cell A1 (row 0, column 0)
            cells.ImportCustomObjects(customers, 0, 0, importOptions);

            // 6. Retrieve and display merged areas after import
            CellArea[] mergedAreas = cells.GetMergedAreas();        // get all merged cells
            Console.WriteLine($"Number of merged areas: {mergedAreas.Length}");
            foreach (CellArea area in mergedAreas)
            {
                Console.WriteLine(
                    $"Merged area - StartRow: {area.StartRow}, StartColumn: {area.StartColumn}, " +
                    $"EndRow: {area.EndRow}, EndColumn: {area.EndColumn}");
            }

            // 7. Save the workbook in XLSX format
            workbook.Save("ImportedDataWithMergedCheck.xlsx", SaveFormat.Xlsx); // save
        }
    }
}