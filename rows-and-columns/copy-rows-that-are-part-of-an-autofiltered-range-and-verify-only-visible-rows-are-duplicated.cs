// Title: Copy Visible Rows from an AutoFiltered Range to Another Worksheet with Aspose.Cells for .NET
// Description: Creates a workbook, populates it with sample data, applies an AutoFilter to show rows where the Category equals "Keep", and then duplicates only the displayed rows to a new sheet using PasteOptions.OnlyVisibleCells. The program prints the copied values and saves the result as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | AutoFilter | visible rows | PasteOptions | OnlyVisibleCells | copy range | filtered data | worksheet duplication | Excel automation
// Common Searches: Aspose.Cells copy filtered rows C# | How to duplicate only visible cells after AutoFilter | PasteOptions OnlyVisibleCells example | Copy AutoFilter results to another sheet Aspose | C# Excel library copy visible rows
// Developer Intent: Duplicate the rows that stay visible after an AutoFilter into a separate worksheet.
// Use Cases: Generate a report that contains only the records matching a filter without altering the original sheet. | Export a filtered view to a new workbook for downstream processing or sharing. | Create a clean snapshot of data for archiving, omitting hidden rows automatically.
// AI Prompts: Show a C# example that copies only the rows displayed by an AutoFilter to a different worksheet using Aspose.Cells. | Explain how to verify that hidden rows are excluded after copying with PasteOptions.OnlyVisibleCells. | Demonstrate dynamic range sizing so the copy works for any number of rows while preserving only visible cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCopyVisibleRowsDemo
{
    // Creates a workbook, populates it with sample data, applies an AutoFilter to show rows where the Category equals "Keep", and then duplicates only the displayed rows to a new sheet using PasteOptions.OnlyVisibleCells. The program prints the copied values and saves the result as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Add a new worksheet and obtain its reference
                int destSheetIndex = workbook.Worksheets.Add();
                Worksheet destSheet = workbook.Worksheets[destSheetIndex];

                // Populate source sheet with sample data (header + 10 rows)
                Cells srcCells = sourceSheet.Cells;
                srcCells["A1"].PutValue("Category");
                srcCells["B1"].PutValue("Value");
                for (int i = 2; i <= 11; i++)
                {
                    // Alternate categories to allow filtering
                    string category = (i % 2 == 0) ? "Keep" : "Skip";
                    srcCells[$"A{i}"].PutValue(category);
                    srcCells[$"B{i}"].PutValue(i * 10);
                }

                // Apply AutoFilter to the header row covering columns A and B
                sourceSheet.AutoFilter.Range = "A1:B11";

                // Filter to show only rows where Category = "Keep"
                sourceSheet.AutoFilter.AddFilter(0, "Keep");
                sourceSheet.AutoFilter.Refresh();

                // Determine the used range of the source data (including hidden rows)
                int totalRows = srcCells.MaxDisplayRange.RowCount;      // includes header
                int totalCols = srcCells.MaxDisplayRange.ColumnCount;   // should be 2

                // Create source and destination ranges covering the same size
                AsposeRange sourceRange = srcCells.CreateRange(0, 0, totalRows, totalCols);
                Cells destCells = destSheet.Cells;
                AsposeRange destRange = destCells.CreateRange(0, 0, totalRows, totalCols);

                // Set paste options to copy only visible cells (i.e., visible rows after filter)
                PasteOptions pasteOptions = new PasteOptions
                {
                    OnlyVisibleCells = true
                };

                // Perform the copy
                destRange.Copy(sourceRange, pasteOptions);

                // Verify the result: print values from destination sheet
                Console.WriteLine("Destination sheet values after copying visible rows:");
                for (int row = 0; row < totalRows; row++)
                {
                    string cat = destCells[row, 0].StringValue;
                    string val = destCells[row, 1].StringValue;

                    // Skip completely empty rows
                    if (string.IsNullOrEmpty(cat) && string.IsNullOrEmpty(val))
                        continue;

                    Console.WriteLine($"Row {row + 1}: Category = {cat}, Value = {val}");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("CopyVisibleRowsResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
