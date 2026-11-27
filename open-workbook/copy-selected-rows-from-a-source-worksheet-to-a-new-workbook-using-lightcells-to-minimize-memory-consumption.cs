// Title: Copy specific rows from an XLSX worksheet to a new workbook using Aspose.Cells LightCells in C# with low memory usage
// Common Searches: Aspose.Cells C# copy selected rows to a new workbook with MemoryPreference | How to use LightCells to copy only certain rows from one Excel file to another in .NET | Copy rows with styles from source.xlsx to destination.xlsx using Aspose.Cells without loading the full workbook | C# example for selective row export with minimal memory using Aspose.Cells LoadOptions | Insert missing rows while copying cells between workbooks with Aspose.Cells LightCells
// Tags: lightcells row extraction | memory‑optimized row transfer Aspose.Cells | transfer rows between XLSX workbooks C# | loadoptions memorypreference Aspose.Cells | preserve cell style during row copy

using System;
using System.IO;
using Aspose.Cells;

// The example loads source.xlsx with LoadOptions.MemoryPreference to reduce memory consumption, selects specific zero‑based rows (e.g., 0, 2, 4), inserts missing rows in a new workbook, copies each cell's value and style using a LightCells‑compatible approach, and saves the result as destination.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "destination.xlsx";

            // Verify source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook with memory‑optimized settings.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                MemorySetting = MemorySetting.MemoryPreference // minimizes memory usage
            };
            var srcWorkbook = new Workbook(sourcePath, loadOptions);
            var srcWorksheet = srcWorkbook.Worksheets[0];

            // Create a new workbook that will hold the copied rows.
            var destWorkbook = new Workbook();
            var destWorksheet = destWorkbook.Worksheets[0];

            // Define the zero‑based indices of the rows you want to copy.
            // Example: copy rows 0, 2 and 4 from the source sheet.
            int[] rowsToCopy = new int[] { 0, 2, 4 };

            // Determine the maximum column index used in the source sheet.
            int maxColumn = srcWorksheet.Cells.MaxColumn;

            // Iterate over each requested row.
            foreach (int srcRowIndex in rowsToCopy)
            {
                // Ensure the source row exists.
                var srcRow = srcWorksheet.Cells.Rows[srcRowIndex];
                if (srcRow == null) continue;

                // Ensure the destination row exists; insert if missing.
                if (destWorksheet.Cells.Rows[srcRowIndex] == null)
                {
                    // Insert a new row at the required index.
                    destWorksheet.Cells.InsertRows(srcRowIndex, 1);
                }

                // Copy each cell in the row.
                for (int col = 0; col <= maxColumn; col++)
                {
                    var srcCell = srcWorksheet.Cells[srcRowIndex, col];
                    if (srcCell == null || srcCell.Type == CellValueType.IsNull) continue;

                    var destCell = destWorksheet.Cells[srcRowIndex, col];

                    // Copy the cell value.
                    destCell.PutValue(srcCell.Value);

                    // Copy the cell style.
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Save the new workbook.
            destWorkbook.Save(destPath);
            Console.WriteLine($"Rows copied successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
