// Title: Aspose.Cells .NET – Keep Header Row Static While Merging Data Rows (ShiftFirstRowDown)
// Description: Demonstrates how to import a DataTable into an Excel workbook with column names as a header, set ImportTableOptions.ShiftFirstRowDown = false (the noadd parameter) to keep the header row fixed, merge the data rows, optionally repeat the header on printed pages with PrintTitleRows, and save the file as XLSX.
// Keywords: Aspose.Cells | ImportTableOptions | ShiftFirstRowDown | static header | merge rows | noadd parameter | PrintTitleRows | C# Excel export | .NET Excel merging
// Common Searches: Aspose.Cells keep header row static when merging | ImportTableOptions ShiftFirstRowDown example C# | prevent header shift during Excel merge Aspose | set PrintTitleRows Aspose.Cells .NET | noadd parameter Aspose.Cells
// Developer Intent: The developer needs to preserve the first (header) row in place while merging subsequent data rows in an Excel sheet generated with Aspose.Cells for .NET.
// Use Cases: Create reports where the header row remains unchanged after merging data rows. | Generate printable worksheets with a repeating header using PrintTitleRows. | Programmatically import tabular data and apply custom merging without disturbing column titles.
// AI Prompts: Explain how ImportTableOptions.ShiftFirstRowDown = false keeps the header static during cell merging in Aspose.Cells for .NET. | Provide a C# example that imports a DataTable, merges data rows, and sets PrintTitleRows to repeat the header. | Describe the impact of the noadd parameter on header positioning when using Aspose.Cells ImportTableOptions.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHeaderStaticDemo
{
    // Demonstrates how to import a DataTable into an Excel workbook with column names as a header, set ImportTableOptions.ShiftFirstRowDown = false (the noadd parameter) to keep the header row fixed, merge the data rows, optionally repeat the header on printed pages with PrintTitleRows, and save the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare a DataTable with a header row and some data rows
            DataTable table = new DataTable();
            table.Columns.Add("Product");
            table.Columns.Add("Quantity");
            table.Rows.Add("Apple", 10);
            table.Rows.Add("Banana", 20);
            table.Rows.Add("Cherry", 30);

            // Import the DataTable starting at cell A1.
            // ShiftFirstRowDown = false ensures the first row (header) stays at its original position.
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // import the column names as header
                ShiftFirstRowDown = false  // noadd parameter: keep header static during merging
            };
            cells.ImportData(table, 0, 0, importOptions);

            // Merge the data rows (rows 2-4) into a single cell to demonstrate merging.
            // The header row (row 1) remains unchanged because of the above option.
            cells.Merge(1, 0, 3, 1); // merges cells A2:B4

            // Optionally, set PrintTitleRows so the header repeats on each printed page
            sheet.PageSetup.PrintTitleRows = "$1:$1";

            // Save the workbook
            workbook.Save("HeaderStaticDuringMerging.xlsx", SaveFormat.Xlsx);
        }
    }
}
