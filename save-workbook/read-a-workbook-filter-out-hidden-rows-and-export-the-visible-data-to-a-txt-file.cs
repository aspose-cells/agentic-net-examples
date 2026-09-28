// Title: Export only visible rows from an Excel workbook to a tab‑delimited TXT file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, uses ExportTableOptions.PlotVisibleRows to skip hidden rows, and saves the visible data as a tab‑separated .txt file. | Show how to export a worksheet's visible range to a DataTable with column headers and then write that DataTable to a tab‑delimited text file in C#. | Demonstrate robust error handling for missing input files and runtime exceptions while exporting visible rows to a TXT file with Aspose.Cells.
// Common Searches: Aspose.Cells C# export visible rows to tab delimited text file | ignore hidden rows when converting Excel to txt using Aspose.Cells | ExportTableOptions PlotVisibleRows usage example .NET | write DataTable to tab delimited file in C# after Excel export | check workbook existence before exporting with Aspose.Cells
// Tags: Aspose.Cells ExportTableOptions PlotVisibleRows | export visible rows to TXT Aspose.Cells | C# tab delimited text export from Excel | DataTable to text file Aspose.Cells | handle missing workbook Aspose.Cells C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportVisibleRows
{
    // Loads an Excel workbook, extracts only rows that are not hidden using ExportTableOptions, and writes the data with column headers to a tab‑separated .txt file, including file existence checks and exception handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the used range size
                int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
                int totalCols = cells.MaxDataColumn + 1;   // MaxDataColumn is zero‑based

                // Set up export options to include only visible rows and column headers
                ExportTableOptions exportOptions = new ExportTableOptions
                {
                    PlotVisibleRows = true,      // Export only rows that are not hidden
                    ExportColumnName = true      // Include column headers in the export
                };

                // Export the data range to a DataTable
                DataTable dataTable = cells.ExportDataTable(0, 0, totalRows, totalCols, exportOptions);

                // Write the DataTable to a TXT file (tab‑separated)
                string outputPath = "visible_data.txt";
                using (StreamWriter writer = new StreamWriter(outputPath))
                {
                    // Write column headers
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        writer.Write(dataTable.Columns[col].ColumnName);
                        if (col < dataTable.Columns.Count - 1)
                            writer.Write('\t');
                    }
                    writer.WriteLine();

                    // Write each data row
                    foreach (DataRow row in dataTable.Rows)
                    {
                        for (int col = 0; col < dataTable.Columns.Count; col++)
                        {
                            writer.Write(row[col]?.ToString() ?? string.Empty);
                            if (col < dataTable.Columns.Count - 1)
                                writer.Write('\t');
                        }
                        writer.WriteLine();
                    }
                }

                Console.WriteLine($"Visible rows exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
