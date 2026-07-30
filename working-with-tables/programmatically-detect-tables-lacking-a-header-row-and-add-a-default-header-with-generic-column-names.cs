// Title: Detect Tables Missing Headers and Insert Default Column Names with Aspose.Cells (C#/.NET)
// Description: A C#/.NET example that opens or creates an Excel workbook, scans every worksheet for ListObject tables without a header row, inserts a top row filled with generic names (Column1, Column2, …), rebuilds the table with the header flag enabled, and saves the updated file. Header presence is inferred by checking the first row for string values.
// Keywords: Aspose.Cells | C# | .NET | Excel ListObject | detect missing table header | add default header row | generic column names | programmatic table header | workbook automation | Excel table without header | US developers | European developers
// Common Searches: How to add a header row to an Aspose.Cells ListObject | Detect tables without headers using Aspose.Cells C# | Add default column names to Excel tables programmatically | Aspose.Cells replace table after inserting header | C# code to insert generic column headers in Excel | Aspose.Cells table header detection heuristic
// Developer Intent: Automatically add a default header row to any Excel table that lacks one.
// Use Cases: Standardize imported spreadsheets before data analysis by ensuring every table has column names. | Prepare workbooks for BI or reporting systems that require explicit headers. | Clean up user‑generated Excel files that contain tables without headers while preserving formatting. | Automate bulk workbook cleanup in enterprise environments.
// AI Prompts: Write C# code using Aspose.Cells to scan all worksheets, find ListObjects without a header row, insert a header row with Column1, Column2… and update the table definition. | Suggest an alternative method to detect missing table headers in Aspose.Cells that does not rely on cell type heuristics. | Explain how to retain existing table styles and formatting when adding a new header row with Aspose.Cells. | Provide a PowerShell script that invokes the compiled .NET assembly to process multiple workbooks for missing table headers.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // A C#/.NET example that opens or creates an Excel workbook, scans every worksheet for ListObject tables without a header row, inserts a top row filled with generic names (Column1, Column2, …), rebuilds the table with the header flag enabled, and saves the updated file. Header presence is inferred by checking the first row for string values.
    class DetectAndAddTableHeaders
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Sample data: create a worksheet with a table that has no header row
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue(1);
                ws.Cells["B1"].PutValue(2);
                ws.Cells["A2"].PutValue(3);
                ws.Cells["B2"].PutValue(4);

                // Add a ListObject (table) without headers (hasHeaders = false)
                int tableIndex = ws.ListObjects.Add(0, 0, 1, 1, false);
                ListObject table = ws.ListObjects[tableIndex];
                table.DisplayName = "DataTable";

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Work on a copy of the ListObjects collection because we may modify it inside the loop
                    ListObject[] tables = new ListObject[sheet.ListObjects.Count];
                    sheet.ListObjects.CopyTo(tables, 0);

                    foreach (ListObject lo in tables)
                    {
                        // Determine if the table already has a header row.
                        // Aspose.Cells ListObject does not expose a direct property for this in older versions,
                        // so we infer it by checking the first row of the table for string values.
                        bool hasHeader = false;
                        int startRow = lo.StartRow;
                        int startCol = lo.StartColumn;
                        int endRow = lo.EndRow;
                        int endCol = lo.EndColumn;

                        // Simple heuristic: if any cell in the first row of the table contains a string, treat it as a header.
                        for (int c = startCol; c <= endCol; c++)
                        {
                            if (sheet.Cells[startRow, c].Type == CellValueType.IsString)
                            {
                                hasHeader = true;
                                break;
                            }
                        }

                        if (!hasHeader)
                        {
                            // Insert a new row at the top of the table to serve as header
                            sheet.Cells.InsertRows(startRow, 1);

                            // Fill the inserted row with generic column names: Column1, Column2, ...
                            int colCount = endCol - startCol + 1;
                            for (int c = 0; c < colCount; c++)
                            {
                                sheet.Cells[startRow, startCol + c].PutValue($"Column{c + 1}");
                            }

                            // Remove the old table definition using its index
                            int loIndex = sheet.ListObjects.IndexOf(lo);
                            if (loIndex >= 0)
                            {
                                sheet.ListObjects.RemoveAt(loIndex);
                            }

                            // Re‑create the table with the new range and indicate that it now has headers
                            int newTableIndex = sheet.ListObjects.Add(startRow, startCol, endRow - startRow + 2, colCount, true);
                            ListObject newTable = sheet.ListObjects[newTableIndex];
                            newTable.DisplayName = lo.DisplayName;

                            // Update column names to match the header cells
                            newTable.UpdateColumnName();
                        }
                    }
                }

                // Save the workbook with the updated tables
                string outputPath = "OutputWithHeaders.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
