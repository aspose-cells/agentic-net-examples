using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook (lifecycle: create and load)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet's cells collection
        Cells cells = workbook.Worksheets[0].Cells;

        // Configure export options to preserve formatting and hierarchy
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            // Export the first row as column names (if present)
            ExportColumnName = true,
            // Export values as strings to keep formatting (e.g., dates, numbers with format)
            ExportAsString = true,
            // Use the cell's style for formatting during export
            FormatStrategy = CellValueFormatStrategy.CellStyle
        };

        // Determine the used range of the worksheet
        int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
        int totalColumns = cells.MaxDataColumn + 1;

        // Export the data (including formatting) to a DataTable
        DataTable dataTable = cells.ExportDataTable(
            firstRow: 0,
            firstColumn: 0,
            totalRows: totalRows,
            totalColumns: totalColumns,
            options: exportOptions);

        // Example usage: print the exported data to the console
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object item in row.ItemArray)
            {
                Console.Write(item + "\t");
            }
            Console.WriteLine();
        }

        // Save the workbook if any modifications were made (lifecycle: save)
        workbook.Save("output.xlsx");
    }
}