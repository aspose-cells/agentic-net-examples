using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCheckMergedCellsDemo
{
    // Sample custom object to be imported
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; } = string.Empty; // initialized to avoid nullable warning
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare sample data
                List<Customer> customers = new List<Customer>
                {
                    new Customer { CustomerId = 1, Name = "Alice" },
                    new Customer { CustomerId = 2, Name = "Bob" },
                    new Customer { CustomerId = 3, Name = "Charlie" }
                };

                // Configure import options with CheckMergedCells enabled
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    CheckMergedCells = true,   // Ensure merged cells are checked during import
                    InsertRows = true,         // Add rows if needed
                    IsFieldNameShown = false   // Do not import property names as header
                };

                // Import the custom objects starting at cell A1 (row 0, column 0)
                // Cast to non‑generic ICollection as required by the API
                worksheet.Cells.ImportCustomObjects((ICollection)customers, 0, 0, importOptions);

                // Save the workbook to a file
                string outputPath = "CheckMergedCellsOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}