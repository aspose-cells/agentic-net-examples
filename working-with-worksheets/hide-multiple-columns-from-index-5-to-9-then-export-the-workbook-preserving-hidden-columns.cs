// Title: C# – Hide Columns 5‑9 and Export Workbook while Preserving Hidden Columns with Aspose.Cells
// Description: Creates a workbook, fills a 10 × 12 range, hides columns 5‑9 (F‑J) using Cells.HideColumns, sets ExportTableOptions.PlotVisibleColumns to false, exports the range to a DataTable that includes hidden columns, prints the data, and saves the file with the hidden columns retained.
// Keywords: Aspose.Cells hide columns | C# hide multiple columns | ExportDataTable hidden columns | ExportTableOptions PlotVisibleColumns | Aspose.Cells workbook export | HideColumns method .NET
// Common Searches: Aspose.Cells hide columns 5 to 9 | ExportDataTable include hidden columns C# | How to keep hidden columns when exporting with Aspose.Cells | C# hide column range and export to DataTable | Aspose.Cells ExportTableOptions false PlotVisibleColumns
// Developer Intent: Hide columns 5‑9 in a worksheet and export the data so that hidden columns are included in the output.
// Use Cases: Generate a report where calculation columns are hidden in Excel but required for backend processing. | Create a template with hidden helper columns, export the full dataset to a database, and keep the hidden columns in the saved file. | Provide end‑users a workbook with concealed columns while still extracting the complete data for analytics.
// AI Prompts: Write C# code with Aspose.Cells to hide columns F‑J and export the entire range to a DataTable, ensuring hidden columns are included. | Show how to configure ExportTableOptions so ExportDataTable exports hidden columns and adds column headers. | Explain the impact of ExportTableOptions.PlotVisibleColumns on the visibility of columns in ExportDataTable results.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills a 10 × 12 range, hides columns 5‑9 (F‑J) using Cells.HideColumns, sets ExportTableOptions.PlotVisibleColumns to false, exports the range to a DataTable that includes hidden columns, prints the data, and saves the file with the hidden columns retained.
    public class HideColumnsAndExport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (10 rows x 12 columns)
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 12; col++)
                    {
                        cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Hide columns with zero‑based indexes 5 through 9 (5 columns total)
                int startColumn = 5;      // Column F
                int columnCount = 5;      // Columns F, G, H, I, J
                cells.HideColumns(startColumn, columnCount);

                // Prepare export options – keep hidden columns in the export
                ExportTableOptions exportOptions = new ExportTableOptions
                {
                    ExportColumnName = true,          // include header row
                    PlotVisibleColumns = false        // false => export hidden columns as well
                };

                // Export the range (first 10 rows, first 12 columns) to a DataTable
                DataTable exportedTable = cells.ExportDataTable(0, 0, 10, 12, exportOptions);

                // Display exported data to verify hidden columns are present
                Console.WriteLine("Exported DataTable (including hidden columns):");
                foreach (DataColumn col in exportedTable.Columns)
                {
                    Console.Write($"{col.ColumnName}\t");
                }
                Console.WriteLine();

                foreach (DataRow row in exportedTable.Rows)
                {
                    foreach (var item in row.ItemArray)
                    {
                        Console.Write($"{item}\t");
                    }
                    Console.WriteLine();
                }

                // Save the workbook – hidden columns are retained in the file
                workbook.Save("HiddenColumnsWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideColumnsAndExport.Run();
        }
    }
}
