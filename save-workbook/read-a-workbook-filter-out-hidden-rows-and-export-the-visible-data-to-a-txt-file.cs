// Title: C# – Export Visible Excel Rows to a Tab‑Delimited TXT with Aspose.Cells
// Description: Loads an Excel workbook, selects the first worksheet, and uses Aspose.Cells ExportTableOptions (PlotVisibleRows = true, ExportColumnName = true) to extract only non‑hidden rows into a DataTable. The DataTable is then written to a TXT file with tab delimiters, preserving column headers.
// Keywords: Aspose.Cells export visible rows | C# ExportTableOptions PlotVisibleRows | save Excel as tab delimited text | .NET export hidden rows | ExportDataTable to TXT | Aspose.Cells visible data export
// Common Searches: Aspose.Cells export only visible rows to txt | C# ExportTableOptions PlotVisibleRows example | How to write Excel visible rows to a tab‑delimited file | ExportDataTable visible rows Aspose.Cells .NET | Save Excel worksheet as text file excluding hidden rows
// Developer Intent: Extract non‑hidden rows from an Excel worksheet and save them as a tab‑separated TXT file using Aspose.Cells for .NET.
// Use Cases: Generate lightweight reports that omit hidden rows before sharing with partners. | Feed visible spreadsheet data into analytics tools that accept tab‑delimited text. | Automate batch extraction of filtered Excel data for downstream processing pipelines.
// AI Prompts: Write C# code with Aspose.Cells to export only visible rows of a worksheet to a CSV file using a custom delimiter. | Explain how ExportTableOptions.PlotVisibleRows works and how to combine it with ExportColumnName to include headers. | Suggest performance‑optimized approaches for exporting large worksheets' visible rows to a text file.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, selects the first worksheet, and uses Aspose.Cells ExportTableOptions (PlotVisibleRows = true, ExportColumnName = true) to extract only non‑hidden rows into a DataTable. The DataTable is then written to a TXT file with tab delimiters, preserving column headers.
class ExportVisibleRowsToTxt
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure export options to include only visible rows
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            PlotVisibleRows = true,      // Export only rows that are not hidden
            ExportColumnName = true      // Include column headers in the export
        };

        // Determine the range to export (entire used range)
        int totalRows = cells.MaxDataRow + 1;       // MaxDataRow is zero‑based
        int totalColumns = cells.MaxDataColumn + 1; // MaxDataColumn is zero‑based

        // Export the visible data to a DataTable
        DataTable dataTable = cells.ExportDataTable(0, 0, totalRows, totalColumns, exportOptions);

        // Write the DataTable to a TXT file (tab‑separated)
        using (StreamWriter writer = new StreamWriter("output.txt"))
        {
            // Write column headers
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                writer.Write(dataTable.Columns[col].ColumnName);
                if (col < dataTable.Columns.Count - 1)
                    writer.Write('\t');
            }
            writer.WriteLine();

            // Write each row of visible data
            foreach (DataRow row in dataTable.Rows)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    writer.Write(row[col]?.ToString());
                    if (col < dataTable.Columns.Count - 1)
                        writer.Write('\t');
                }
                writer.WriteLine();
            }
        }

        Console.WriteLine("Visible rows have been exported to output.txt");
    }
}
