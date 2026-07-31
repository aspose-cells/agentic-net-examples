// Title: Skip Merged Cells When Importing a DataTable with Aspose.Cells (CheckMergedCells = false)
// Description: C# example that creates a workbook, merges cells D4:D5, builds a DataTable, and imports it using ImportTableOptions. By setting CheckMergedCells to false and InsertRows to true, the import skips the merged range, preserves existing merged content, and adds new rows. The workbook is saved as ImportSkipMergedCells.xlsx.
// Keywords: Aspose.Cells ImportTableOptions | CheckMergedCells false | ImportData DataTable merged cells | C# Aspose.Cells skip merged cells | InsertRows true | Excel merged cell import | Aspose.Cells .NET example
// Common Searches: Aspose.Cells import DataTable without overwriting merged cells | ImportTableOptions.CheckMergedCells false example | How to ignore merged ranges when using ImportData | C# Aspose.Cells skip merged cells during import | InsertRows true with ImportTableOptions
// Developer Intent: Import a DataTable into a worksheet while leaving existing merged cells untouched.
// Use Cases: Populate a template that contains merged header cells without destroying the header layout. | Add transactional rows to a sheet that has merged summary sections, preserving formatting. | Generate invoices where merged cells define sections, and external data must be inserted without altering those sections.
// AI Prompts: Show a C# code snippet that uses Aspose.Cells ImportTableOptions to import a DataTable while skipping merged cells. | Explain how CheckMergedCells = false and InsertRows = true affect data import into a worksheet with merged ranges. | Provide step‑by‑step guidance for preserving merged cells when importing external data into an Aspose.Cells workbook.

using System;
using System.Data;
using Aspose.Cells;

// C# example that creates a workbook, merges cells D4:D5, builds a DataTable, and imports it using ImportTableOptions. By setting CheckMergedCells to false and InsertRows to true, the import skips the merged range, preserves existing merged content, and adds new rows. The workbook is saved as ImportSkipMergedCells.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a merged cell range D4:D5 (row index 3, column index 3)
        cells.Merge(3, 3, 2, 1);
        cells[3, 3].PutValue("MergedValue");

        // Prepare a DataTable with sample data to import
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Rows.Add(1, "Alice");
        dataTable.Rows.Add(2, "Bob");

        // Set import options and disable checking merged cells
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = true,   // import column headers
            InsertRows = true,         // insert rows instead of overwriting
            CheckMergedCells = false   // skip writing into merged cell ranges
        };

        // Import the data starting at the same location as the merged cells
        // Because CheckMergedCells is false, the merged cells will be left unchanged
        cells.ImportData(dataTable, 3, 3, importOptions);

        // Save the workbook to a file
        workbook.Save("ImportSkipMergedCells.xlsx", SaveFormat.Xlsx);
    }
}
