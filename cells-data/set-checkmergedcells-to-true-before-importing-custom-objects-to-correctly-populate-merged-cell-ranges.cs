// Title: C# – Preserve Merged Cells When Importing Custom Objects with Aspose.Cells (CheckMergedCells = true)
// Description: Shows how to enable ImportTableOptions.CheckMergedCells before calling Worksheet.Cells.ImportCustomObjects so that a pre‑merged range (e.g., D4:D5) stays intact after importing a List<Customer>. The sample creates a workbook, merges cells, configures import options, imports data, saves the file, and confirms the merge.
// Keywords: Aspose.Cells | ImportCustomObjects | CheckMergedCells | merged cells | C# .NET | ImportTableOptions | custom object list | Excel automation | preserve merged ranges | worksheet import | Aspose.Cells example | GitHub Aspose.Cells | Excel template merging
// Common Searches: Aspose.Cells keep merged cells after ImportCustomObjects | ImportCustomObjects CheckMergedCells true C# | How to preserve merged ranges when importing data with Aspose.Cells | Import list of objects into Excel without breaking merges | ImportTableOptions.CheckMergedCells usage example
// Developer Intent: Enable the CheckMergedCells flag so existing merged cell blocks are not altered when custom objects are imported into a worksheet.
// Use Cases: Import a collection of business objects beneath a merged title row in a reporting template. | Populate data in a pre‑formatted Excel sheet that uses merged cells for section headings. | Automate data entry into a workbook while maintaining layout integrity defined by merged ranges.
// AI Prompts: Write C# code that imports a List<Customer> into an Aspose.Cells worksheet and keeps any merged cells by setting ImportTableOptions.CheckMergedCells to true. | Provide an example of using ImportTableOptions with CheckMergedCells enabled to import data into a workbook containing merged header cells. | Explain how to verify that merged cells remain after ImportCustomObjects and suggest error‑handling strategies.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Sample custom object to be imported
    // Shows how to enable ImportTableOptions.CheckMergedCells before calling Worksheet.Cells.ImportCustomObjects so that a pre‑merged range (e.g., D4:D5) stays intact after importing a List<Customer>. The sample creates a workbook, merges cells, configures import options, imports data, saves the file, and confirms the merge.
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; } = string.Empty; // initialize to avoid nullable warning
    }

    public class ImportCustomObjectsWithMergedCellsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a merged cell range for demonstration (e.g., D4:D5)
                // Row and column indexes are zero‑based, so D4 is (3,3)
                worksheet.Cells.Merge(3, 3, 2, 1);
                worksheet.Cells[3, 3].PutValue("MergedValue");

                // Prepare a list of custom objects to import
                List<Customer> customers = new List<Customer>
                {
                    new Customer { CustomerId = 1, Name = "Alice" },
                    new Customer { CustomerId = 2, Name = "Bob" },
                    new Customer { CustomerId = 3, Name = "Charlie" }
                };

                // Configure import options and enable merged‑cell checking
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    // Do not import field names as a header row
                    IsFieldNameShown = false,
                    // Insert rows if needed to accommodate data
                    InsertRows = true,
                    // Preserve merged cells during import
                    CheckMergedCells = true
                };

                // Import the custom objects starting at cell A1 (row 0, column 0)
                worksheet.Cells.ImportCustomObjects((ICollection)customers, 0, 0, importOptions);

                // Save the workbook
                string outputPath = "ImportCustomObjectsWithMergedCellsDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Verify that the merged cell is still present after import
                Cell mergedCell = worksheet.Cells["D4"];
                Console.WriteLine($"Merged cell value: {mergedCell.StringValue}");
                Console.WriteLine($"Is cell merged: {mergedCell.IsMerged}");
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            ImportCustomObjectsWithMergedCellsDemo.Run();
        }
    }
}
