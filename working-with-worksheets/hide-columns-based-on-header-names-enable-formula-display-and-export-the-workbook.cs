// Title: C# – Hide Columns by Header, Show Formulas, Export Visible Data with Aspose.Cells
// Description: Creates a workbook, hides any column whose header equals "Secret", enables ShowFormulas to display formulas, adds a SUM formula, exports only visible columns (with headers) to a DataTable via ExportTableOptions, prints the table, and saves the file as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | hide column by header | ShowFormulas | ExportDataTable | ExportTableOptions | PlotVisibleColumns | visible columns export | DataTable export Aspose | save workbook XLSX | Aspose.Cells example | GitHub sample code | Aspose.Cells .NET
// Common Searches: Aspose.Cells hide column with specific header | Export only visible columns to DataTable Aspose.Cells | Show formulas instead of values in Aspose.Cells workbook | ExportTableOptions PlotVisibleColumns example | C# Aspose.Cells hide column and save XLSX
// Developer Intent: Hide columns matching a given header, display formulas in the sheet, export the visible portion to a DataTable, and persist the workbook.
// Use Cases: Prepare reports that exclude confidential columns before distribution. | Generate a DataTable containing only user‑visible fields for further processing or UI binding. | Debug complex worksheets by toggling ShowFormulas to view underlying formulas. | Automate Excel file creation where hidden columns must not appear in exported data.
// AI Prompts: Write C# code with Aspose.Cells that hides columns whose header matches a supplied string and exports only the visible columns to a DataTable. | Show how to enable ShowFormulas, add a SUM formula, and save the workbook as an XLSX file using Aspose.Cells. | Explain the impact of ExportTableOptions.PlotVisibleColumns and ExportColumnName on the DataTable returned by ExportDataTable.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, hides any column whose header equals "Secret", enables ShowFormulas to display formulas, adds a SUM formula, exports only visible columns (with headers) to a DataTable via ExportTableOptions, prints the table, and saves the file as XLSX using Aspose.Cells for .NET.
    public class HideColumnsAndExportDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header row
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Secret");   // This column will be hidden
                sheet.Cells["D1"].PutValue("Amount");

                // Populate some data rows
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue("TopSecret");
                sheet.Cells["D2"].PutValue(100);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue("Classified");
                sheet.Cells["D3"].PutValue(200);

                // Hide columns whose header equals "Secret"
                for (int col = 0; col <= sheet.Cells.MaxColumn; col++)
                {
                    Cell headerCell = sheet.Cells[0, col];
                    if (headerCell != null && headerCell.Type == CellValueType.IsString)
                    {
                        if (headerCell.StringValue.Equals("Secret", StringComparison.OrdinalIgnoreCase))
                        {
                            // Hide the column (rule: HideColumn)
                            sheet.Cells.HideColumn(col);
                        }
                    }
                }

                // Enable formula display (show formulas instead of results)
                sheet.ShowFormulas = true;

                // Example formula (will be shown because ShowFormulas = true)
                sheet.Cells["E2"].Formula = "=SUM(D2:D3)";

                // Export visible data to a DataTable, exporting only visible columns
                ExportTableOptions exportOptions = new ExportTableOptions
                {
                    PlotVisibleColumns = true,   // Export only visible columns
                    ExportColumnName = true      // Include header names as column names
                };

                // Determine used range
                int totalRows = sheet.Cells.MaxDataRow + 1;   // +1 because rows are zero‑based
                int totalCols = sheet.Cells.MaxDataColumn + 1;

                // Export the data
                DataTable exportedTable = sheet.Cells.ExportDataTable(0, 0, totalRows, totalCols, exportOptions);

                // Output exported DataTable to console
                Console.WriteLine("Exported DataTable (visible columns only):");
                foreach (DataColumn col in exportedTable.Columns)
                {
                    Console.Write(col.ColumnName + "\t");
                }
                Console.WriteLine();
                foreach (DataRow row in exportedTable.Rows)
                {
                    foreach (var item in row.ItemArray)
                    {
                        Console.Write(item + "\t");
                    }
                    Console.WriteLine();
                }

                // Save the workbook to an XLSX file (lifecycle rule: save)
                string outputPath = "HiddenColumnsDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                HideColumnsAndExportDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
