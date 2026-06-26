using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportVisibleRows
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set up export options to include only visible rows and columns
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                PlotVisibleRows = true,      // Export only rows that are not hidden
                PlotVisibleColumns = true,   // Export only columns that are not hidden
                ExportColumnName = true      // Include column headers in the export
            };

            // Determine the range to export (entire used range)
            int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
            int totalColumns = cells.MaxDataColumn + 1;

            // Export the visible data to a DataTable
            DataTable dataTable = cells.ExportDataTable(0, 0, totalRows, totalColumns, exportOptions);

            // Write the DataTable to a TXT file (tab‑separated)
            using (StreamWriter writer = new StreamWriter("output.txt"))
            {
                // Write column headers
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    writer.Write(dataTable.Columns[col].ColumnName);
                    if (col < dataTable.Columns.Count - 1) writer.Write('\t');
                }
                writer.WriteLine();

                // Write each data row
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        writer.Write(row[col]?.ToString() ?? string.Empty);
                        if (col < dataTable.Columns.Count - 1) writer.Write('\t');
                    }
                    writer.WriteLine();
                }
            }

            Console.WriteLine("Visible data exported to output.txt");
        }
    }
}