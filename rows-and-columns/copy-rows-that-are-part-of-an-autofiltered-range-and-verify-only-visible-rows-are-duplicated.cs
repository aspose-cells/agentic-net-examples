// Title: Copy visible rows from an AutoFiltered range using Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills column A with a header and ten numeric rows, applies an AutoFilter (values > 5), selects only the filtered rows, and copies them to a new location with PasteOptions.OnlyVisibleCells. The code prints the copied values for verification and saves the file.
// Keywords: Aspose.Cells | C# | Copy visible rows | AutoFilter | OnlyVisibleCells | PasteOptions | filtered data copy | Excel automation | duplicate filtered rows | range copy Aspose
// Common Searches: Aspose.Cells copy only visible rows | PasteOptions OnlyVisibleCells C# example | How to duplicate filtered rows with Aspose.Cells | Copy AutoFilter results to another range .NET | Aspose.Cells copy visible cells after filter
// Developer Intent: Duplicate the rows that remain visible after applying an AutoFilter, excluding hidden rows, and paste them to a separate range.
// Use Cases: Create a summary section that lists only rows where a column meets a condition. | Archive filtered records to a new area without altering the original dataset. | Generate a printable report containing only the rows that satisfy the filter criteria.
// AI Prompts: Show a C# snippet that copies only visible cells from an AutoFiltered range using Aspose.Cells. | Explain how to confirm that OnlyVisibleCells copied just the filtered rows. | Provide code to copy visible rows to another worksheet while preserving the original filter.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCopyVisibleRowsDemo
{
    // C# example that creates a workbook, fills column A with a header and ten numeric rows, applies an AutoFilter (values > 5), selects only the filtered rows, and copies them to a new location with PasteOptions.OnlyVisibleCells. The code prints the copied values for verification and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (header + 10 rows) in column A
                cells["A1"].PutValue("Header");
                for (int i = 2; i <= 11; i++) // rows 2..11 (index 1..10)
                {
                    cells[$"A{i}"].PutValue(i - 1); // values 1..10
                }

                // Apply an AutoFilter to the header row covering column A
                worksheet.AutoFilter.Range = "A1:A11";

                // Filter to show only rows where the value is greater than 5
                worksheet.AutoFilter.Custom(0, FilterOperatorType.GreaterThan, 5);
                worksheet.AutoFilter.Refresh(); // hide rows that do not meet the criteria

                // Determine the area the AutoFilter applies to (including header)
                CellArea filterArea = worksheet.AutoFilter.GetCellArea();

                // Define source range: data rows only (exclude header)
                int sourceStartRow = filterArea.StartRow + 1; // first data row
                int sourceStartColumn = filterArea.StartColumn;
                int rowCount = filterArea.EndRow - sourceStartRow + 1;
                int columnCount = filterArea.EndColumn - filterArea.StartColumn + 1;

                AsposeRange sourceRange = cells.CreateRange(sourceStartRow, sourceStartColumn, rowCount, columnCount);

                // Define destination range placed below the original data
                int destStartRow = filterArea.EndRow + 2; // one empty row gap
                AsposeRange destRange = cells.CreateRange(destStartRow, sourceStartColumn, rowCount, columnCount);

                // Set paste options to copy only visible cells (i.e., skip hidden rows)
                PasteOptions pasteOptions = new PasteOptions
                {
                    OnlyVisibleCells = true
                };

                // Perform the copy
                destRange.Copy(sourceRange, pasteOptions);

                // Verify the copy: print values from the destination range
                Console.WriteLine("Copied values (only visible rows should appear):");
                for (int r = 0; r < rowCount; r++)
                {
                    // Destination cell address
                    string address = CellsHelper.CellIndexToName(destStartRow + r, sourceStartColumn);
                    string value = cells[address].StringValue;
                    Console.WriteLine($"{address}: {(string.IsNullOrEmpty(value) ? "<empty>" : value)}");
                }

                // Save the workbook (optional, demonstrates that the file contains the copied rows)
                string outputPath = "CopyVisibleRowsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
