// Title: Extract ListObject Tables from an XLSM Workbook to Separate XLSX Files with Aspose.Cells (C#)
// Description: C# sample that loads a macro‑enabled XLSM workbook, walks through every worksheet, extracts each ListObject (Excel table) into a DataTable, creates a new workbook, writes the headers and rows, and saves the result as an individual XLSX file. Includes file‑existence validation and robust error handling.
// Keywords: Aspose.Cells extract tables | C# ListObject to XLSX | split XLSM into multiple workbooks | export Excel table programmatically | iterate worksheets Aspose.Cells | save each table as separate file | macro enabled workbook processing | Aspose.Cells table extraction example
// Common Searches: how to extract tables from an XLSM using Aspose.Cells C# | save each ListObject as a separate XLSX file | split macro workbook into individual Excel files | Aspose.Cells export ListObject to new workbook | C# code to iterate worksheets and tables in Excel
// Developer Intent: Programmatically separate every table in a macro‑enabled workbook into its own XLSX file.
// Use Cases: Create standalone reports for each data table stored in a shared macro workbook. | Prepare per‑table files for downstream analytics pipelines that require single‑table inputs. | Automate archiving of consolidated workbooks by splitting them into individual, version‑controlled files.
// AI Prompts: Generate C# code that uses Aspose.Cells to extract all ListObjects from an XLSM and save each as a separate XLSX, with file‑existence checks. | Show how to preserve original table styles, column widths, and formatting when exporting tables to new workbooks. | Suggest a naming convention for the output files that includes the source worksheet name and the ListObject name instead of a numeric counter.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

namespace AsposeCellsTableExtractor
{
    // C# sample that loads a macro‑enabled XLSM workbook, walks through every worksheet, extracts each ListObject (Excel table) into a DataTable, creates a new workbook, writes the headers and rows, and saves the result as an individual XLSX file. Includes file‑existence validation and robust error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source XLSM workbook
                string sourcePath = "source.xlsm";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: The file '{sourcePath}' was not found.");
                    return;
                }

                // Load the macro‑enabled workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                int tableCounter = 0;

                // Iterate through all worksheets
                foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                {
                    // Iterate through all tables (ListObjects) in the worksheet
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        try
                        {
                            // Create a new empty workbook for the current table
                            Workbook tableWorkbook = new Workbook();

                            // Get the first (default) worksheet of the new workbook
                            Worksheet tableSheet = tableWorkbook.Worksheets[0];

                            // Calculate the size of the table
                            int rowCount = table.EndRow - table.StartRow + 1;
                            int columnCount = table.EndColumn - table.StartColumn + 1;

                            // Export the table data to a DataTable
                            DataTable dt = sheet.Cells.ExportDataTable(
                                table.StartRow,          // start row of the table
                                table.StartColumn,       // start column of the table
                                rowCount,                // number of rows in the table
                                columnCount,             // number of columns in the table
                                true);                   // include column names as headers

                            // Manually import the DataTable into the new worksheet
                            // Write column headers
                            for (int col = 0; col < dt.Columns.Count; col++)
                            {
                                tableSheet.Cells[0, col].PutValue(dt.Columns[col].ColumnName);
                            }

                            // Write data rows
                            for (int row = 0; row < dt.Rows.Count; row++)
                            {
                                for (int col = 0; col < dt.Columns.Count; col++)
                                {
                                    tableSheet.Cells[row + 1, col].PutValue(dt.Rows[row][col]);
                                }
                            }

                            // Build a file name for the extracted table
                            string outputPath = $"Table_{tableCounter}.xlsx";

                            // Save the new workbook as an XLSX file
                            tableWorkbook.Save(outputPath, SaveFormat.Xlsx);

                            Console.WriteLine($"Saved table {tableCounter} to '{outputPath}'.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to extract table {tableCounter}: {ex.Message}");
                        }

                        tableCounter++;
                    }
                }

                Console.WriteLine("All tables have been processed.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
