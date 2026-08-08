// Title: Skip merged‑cell validation when importing a DataTable with Aspose.Cells (C#)
// Description: Shows how to import a DataTable into an Excel worksheet using Aspose.Cells C# while setting ImportTableOptions.CheckMergedCells = false to overwrite merged ranges.
// Keywords: Aspose.Cells | C# | ImportTableOptions | CheckMergedCells | ImportData | DataTable | merged cells | skip merged validation | Excel export | worksheet import
// Common Searches: Aspose.Cells import DataTable without checking merged cells | Set ImportTableOptions.CheckMergedCells false C# | ImportData skip merged ranges Aspose | Overwrite merged cells Aspose.Cells | ImportTableOptions example C#
// Developer Intent: Load tabular data into an Excel sheet that contains merged cells without raising validation errors, allowing the new data to replace the merged areas.
// Use Cases: Populate a pre‑designed report template where header rows are merged, but the data section must be filled programmatically. | Migrate query results into a worksheet that uses merged cells for layout, ensuring the import proceeds without exceptions. | Automate data refresh for dashboards that rely on merged cells as placeholders, overwriting them with fresh values.
// AI Prompts: Write C# code that uses Aspose.Cells to import a DataTable into a worksheet with ImportTableOptions.CheckMergedCells set to false. | Explain the effect of the CheckMergedCells property on Cells.ImportData and how to safely overwrite merged ranges. | Provide step‑by‑step instructions for importing data into an Excel sheet containing merged cells without triggering validation errors.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportSkipMergedCells
{
    // Shows how to import a DataTable into an Excel worksheet using Aspose.Cells C# while setting ImportTableOptions.CheckMergedCells = false to overwrite merged ranges.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a merged cell range for demonstration (e.g., merge B2:C3)
            cells.Merge(1, 1, 2, 2); // Merges cells B2:C3
            cells[1, 1].PutValue("MergedCell");

            // Prepare sample data to import
            DataTable table = new DataTable("Sample");
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add(1, "Alice");
            table.Rows.Add(2, "Bob");
            table.Rows.Add(3, "Charlie");

            // Configure import options
            ImportTableOptions options = new ImportTableOptions
            {
                // Skip checking merged cells so data will be written over them
                CheckMergedCells = false,
                // Do not import column headers
                IsFieldNameShown = false,
                // Insert rows if needed
                InsertRows = true
            };

            // Import the DataTable starting at cell A1 (row 0, column 0)
            cells.ImportData(table, 0, 0, options);

            // Save the workbook to a file
            workbook.Save("ImportSkipMergedCells.xlsx", SaveFormat.Xlsx);
        }
    }
}
