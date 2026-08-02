// Title: C# – Remove Empty Rows from an Aspose.Cells Worksheet Using a Bottom‑Up Loop
// Description: A concise example that creates a workbook, adds sample data with blank rows, then uses MaxDataRow and MaxDataColumn to scan each row from the bottom up. Rows whose cells are all of type IsNull are deleted with Cells.DeleteRow, and the cleaned workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# delete empty rows | remove blank rows Aspose.Cells | Aspose.Cells MaxDataRow | Aspose.Cells MaxDataColumn | Cells.DeleteRow example | C# Excel remove empty rows | Aspose.Cells worksheet cleanup | bottom up row deletion Aspose.Cells
// Common Searches: how to delete empty rows in Aspose.Cells C# | remove blank rows from Excel using Aspose.Cells | Aspose.Cells delete rows with no data | C# loop to clean up empty rows in worksheet | Aspose.Cells bottom up row deletion to avoid index shift
// Developer Intent: Eliminate every completely empty row from a worksheet programmatically.
// Use Cases: Sanitize imported CSV or database data by stripping rows that contain no values before further processing. | Prepare a template workbook for end‑users by removing placeholder rows left blank during design. | Compress generated reports by deleting trailing blank rows after dynamic data insertion.
// AI Prompts: Show how to modify the loop so rows that contain only formulas are kept while empty rows are removed. | Provide an alternative solution that uses Worksheet.Rows collection to delete empty rows in a single operation. | Explain how to preserve rows with comments or cell formatting even when their values are null.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A concise example that creates a workbook, adds sample data with blank rows, then uses MaxDataRow and MaxDataColumn to scan each row from the bottom up. Rows whose cells are all of type IsNull are deleted with Cells.DeleteRow, and the cleaned workbook is saved as an XLSX file.
    class RemoveEmptyRowsWithLoop
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample data with some blank rows
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue("Data1");
                cells["A4"].PutValue("Data2"); // Row 3 is blank
                cells["A6"].PutValue("Data3"); // Row 5 is blank

                // Determine the last row that contains data
                int lastRow = cells.MaxDataRow;

                // Loop from the bottom up to avoid index shifting after deletion
                for (int row = lastRow; row >= 0; row--)
                {
                    bool isEmpty = true;

                    // Check each column up to the last column that contains data
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        // If any cell in the row is not blank, the row is not empty
                        Cell cell = cells[row, col];
                        if (cell.Type != CellValueType.IsNull)
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    // Delete the row if it is empty
                    if (isEmpty)
                    {
                        cells.DeleteRow(row);
                    }
                }

                // Define output file path
                string outputPath = "RemovedEmptyRows.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
