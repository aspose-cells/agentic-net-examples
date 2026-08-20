// Title: Copy Excel Range to a New Workbook while Preserving Column Widths & Row Heights – Aspose.Cells C#
// Description: Load a source workbook, define a range (e.g., A1:C5), create an empty destination workbook, and use `Range.Copy` to transfer data, formulas, and formatting. Then copy each source column width with `SetColumnWidth` and each row height with `SetRowHeight` so the new file matches the original layout before saving.
// Keywords: Aspose.Cells copy range C# | preserve column width Aspose.Cells | preserve row height Aspose.Cells | copy range to new workbook .NET | Excel range dimensions copy | Range.Copy with formatting | C# Excel export preserving layout
// Common Searches: Aspose.Cells copy range to another workbook keep column width | C# preserve row height when copying Excel cells | How to duplicate a table in a new Excel file with original dimensions | Copy range A1:C5 to new workbook Aspose.Cells .NET | Transfer Excel range preserving layout using Aspose
// Developer Intent: Duplicate a specific cell range from an existing workbook into a fresh workbook while retaining the original column widths, row heights, and formatting.
// Use Cases: Create a standalone report by extracting a table from a master workbook without losing its visual layout. | Split a large worksheet into multiple files, each containing a segment that looks identical to the source. | Automate the generation of printable sheets where exact column and row dimensions are required.
// AI Prompts: Write C# code with Aspose.Cells that copies a range to a new workbook and keeps column widths and row heights. | Explain how to use SetColumnWidth and SetRowHeight after Range.Copy to maintain dimensions in Aspose.Cells. | Provide an example that copies a range with formulas, styles, and merged cells to another workbook while preserving layout.

using Aspose.Cells;
using System;
using System.IO;

// Load a source workbook, define a range (e.g., A1:C5), create an empty destination workbook, and use `Range.Copy` to transfer data, formulas, and formatting. Then copy each source column width with `SetColumnWidth` and each row height with `SetRowHeight` so the new file matches the original layout before saving.
class CopyRangePreserveDimensions
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Ensure source file exists; create a simple workbook if missing
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                // Populate some sample data
                tempSheet.Cells["A1"].PutValue("Header1");
                tempSheet.Cells["B1"].PutValue("Header2");
                tempSheet.Cells["C1"].PutValue("Header3");
                for (int r = 2; r <= 5; r++)
                {
                    tempSheet.Cells[$"A{r}"].PutValue($"R{r - 1}C1");
                    tempSheet.Cells[$"B{r}"].PutValue($"R{r - 1}C2");
                    tempSheet.Cells[$"C{r}"].PutValue($"R{r - 1}C3");
                }
                tempWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(sourcePath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Define the source range to copy
            const string sourceRangeAddress = "A1:C5";
            Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange(sourceRangeAddress);

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Destination start position (top‑left cell where the range will be placed)
            int destStartRow = 0;   // row index (0‑based)
            int destStartCol = 0;   // column index (0‑based)

            // Create a destination range with the same size as the source range
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange(
                destStartRow,
                destStartCol,
                srcRange.RowCount,
                srcRange.ColumnCount);

            // Copy data, formulas, formatting, etc. from source to destination
            destRange.Copy(srcRange);

            // ----- Preserve column widths -----
            for (int i = 0; i < srcRange.ColumnCount; i++)
            {
                int srcColIndex = srcRange.FirstColumn + i;
                double colWidth = srcSheet.Cells.GetColumnWidth(srcColIndex); // width in characters
                int destColIndex = destStartCol + i;
                destSheet.Cells.SetColumnWidth(destColIndex, colWidth);
            }

            // ----- Preserve row heights -----
            for (int i = 0; i < srcRange.RowCount; i++)
            {
                int srcRowIndex = srcRange.FirstRow + i;
                double rowHeight = srcSheet.Cells.GetRowHeight(srcRowIndex); // height in points
                int destRowIndex = destStartRow + i;
                destSheet.Cells.SetRowHeight(destRowIndex, rowHeight);
            }

            // Ensure the directory for the destination file exists
            string destDir = Path.GetDirectoryName(Path.GetFullPath(destinationPath));
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the new workbook
            destWorkbook.Save(destinationPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook copied successfully to '{destinationPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
