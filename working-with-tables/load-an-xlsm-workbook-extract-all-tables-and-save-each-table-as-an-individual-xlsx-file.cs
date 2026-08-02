// Title: Extract ListObject tables from an XLSM workbook to separate XLSX files using Aspose.Cells for .NET
// Description: Loads a macro‑enabled XLSM workbook, iterates through every worksheet, extracts each ListObject (Excel table) and writes its cell values into a new workbook, saving each table as an individual XLSX file (Table_0.xlsx, Table_1.xlsx, …).
// Keywords: Aspose.Cells | C# | XLSM | XLSX | ListObject | extract tables | export Excel tables | macro enabled workbook | save each table | Excel table extraction
// Common Searches: how to export each table from an XLSM file using Aspose.Cells | C# extract ListObject tables to separate workbooks | Aspose.Cells copy table range to new XLSX file | extract tables from macro enabled workbook Aspose.Cells | save Excel tables as individual files C#
// Developer Intent: Extract every table in a macro‑enabled workbook and save each one as its own XLSX file.
// Use Cases: Create standalone reports for each data table inside a shared XLSM workbook. | Feed individual tables into downstream pipelines that only accept plain XLSX files. | Distribute per‑table files to different teams while preserving the original macro workbook.
// AI Prompts: Generate C# code with Aspose.Cells that extracts all ListObjects from an XLSM file and saves each as a separate XLSX workbook, preserving only cell values. | Suggest code changes to copy table formatting, column widths, and styles when extracting tables with Aspose.Cells. | Provide performance‑tuning tips for extracting a large number of tables from a big XLSM workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTableExtractor
{
    // Loads a macro‑enabled XLSM workbook, iterates through every worksheet, extracts each ListObject (Excel table) and writes its cell values into a new workbook, saving each table as an individual XLSX file (Table_0.xlsx, Table_1.xlsx, …).
    class Program
    {
        static void Main()
        {
            // Path to the source XLSM workbook
            string sourcePath = "input.xlsm";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook (lifecycle: load)
                Workbook sourceWorkbook = new Workbook(sourcePath);
                int tableCounter = 0;

                // Iterate through each worksheet in the source workbook
                foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
                {
                    // Iterate through each table (ListObject) in the worksheet
                    foreach (ListObject table in sourceSheet.ListObjects)
                    {
                        try
                        {
                            // Create a new workbook for the individual table (lifecycle: create)
                            Workbook tableWorkbook = new Workbook();

                            // Get the first (and only) worksheet in the new workbook
                            Worksheet destSheet = tableWorkbook.Worksheets[0];

                            // Determine the range of the table
                            AsposeRange range = table.DataRange;
                            int startRow = range.FirstRow;
                            int endRow = startRow + range.RowCount - 1;
                            int startCol = range.FirstColumn;
                            int endCol = startCol + range.ColumnCount - 1;

                            // Copy the table data cell by cell
                            for (int row = startRow; row <= endRow; row++)
                            {
                                for (int col = startCol; col <= endCol; col++)
                                {
                                    // Retrieve the value from the source cell
                                    object value = sourceSheet.Cells[row, col].Value;

                                    // Place the value into the destination workbook, preserving relative positions
                                    destSheet.Cells[row - startRow, col - startCol].PutValue(value);
                                }
                            }

                            // Save the individual table as an XLSX file (lifecycle: save)
                            string outputPath = $"Table_{tableCounter}.xlsx";
                            tableWorkbook.Save(outputPath, SaveFormat.Xlsx);
                            Console.WriteLine($"Saved table {tableCounter} to {outputPath}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing table {tableCounter}: {ex.Message}");
                        }

                        tableCounter++;
                    }
                }

                Console.WriteLine("All tables have been extracted and saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
            }
        }
    }
}
