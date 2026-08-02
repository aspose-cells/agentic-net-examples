// Title: C# – Copy a Range to a New Workbook While Preserving Column Widths and Row Heights with Aspose.Cells
// Description: Shows how to load a source .xlsx file, define a range (e.g., A1:C10), and copy it into a fresh workbook using Aspose.Cells for .NET. The sample leverages CopyOptions.ColumnCharacterWidth together with CopyRows and CopyColumns to retain original column widths, row heights, and cell formatting before saving the result as an Xlsx file.
// Keywords: Aspose.Cells | C# | copy range | new workbook | preserve column width | preserve row height | CopyOptions | CopyRows | CopyColumns | Excel automation | Aspose.Cells .NET | Excel file manipulation | US developers | UK developers
// Common Searches: Aspose.Cells copy range to another workbook preserving column width | How to keep row heights when copying cells with Aspose.Cells C# | Copy A1:C10 to a new Excel file using Aspose.Cells .NET | C# example for copying a range with formatting Aspose.Cells | Preserve Excel layout while extracting a table with Aspose.Cells
// Developer Intent: Transfer a defined block of cells from an existing workbook into a separate workbook without losing layout or formatting.
// Use Cases: Create a standalone template from a formatted table in a master workbook. | Generate a report that reuses a specific data block with exact column and row dimensions. | Split a large spreadsheet into multiple files, each containing a particular range with its original layout intact.
// AI Prompts: Write C# code that copies range A1:D20 from source.xlsx to a new workbook, keeping column widths and row heights using Aspose.Cells. | Explain the role of CopyOptions.ColumnCharacterWidth in Aspose.Cells and give a short code snippet that demonstrates its effect. | Create a reusable C# method accepting source path, range address, and destination path, then copies the range with all formatting and dimensions via Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopy
{
    // Shows how to load a source .xlsx file, define a range (e.g., A1:C10), and copy it into a fresh workbook using Aspose.Cells for .NET. The sample leverages CopyOptions.ColumnCharacterWidth together with CopyRows and CopyColumns to retain original column widths, row heights, and cell formatting before saving the result as an Xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "Source.xlsx";
                const string destinationPath = "Destination.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);
                Worksheet srcSheet = srcWorkbook.Worksheets[0];

                // Define the range to copy (e.g., A1:C10)
                AsposeRange srcRange = srcSheet.Cells.CreateRange("A1:C10");

                // Get range boundaries
                int srcFirstRow = srcRange.FirstRow;          // 0‑based index
                int srcFirstColumn = srcRange.FirstColumn;    // 0‑based index
                int rowCount = srcRange.RowCount;
                int columnCount = srcRange.ColumnCount;

                // Create a new (empty) destination workbook
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Configure copy options to preserve column widths (in character units)
                CopyOptions copyOptions = new CopyOptions
                {
                    ColumnCharacterWidth = true   // ensures column widths are copied
                };

                // -----------------------------------------------------------------
                // 1. Copy rows (including row heights and cell data/formats)
                //    Destination rows start at index 0.
                // -----------------------------------------------------------------
                destSheet.Cells.CopyRows(
                    srcSheet.Cells,          // source cells
                    srcFirstRow,            // source start row
                    0,                      // destination start row
                    rowCount,               // number of rows to copy
                    copyOptions);           // copy options (preserve column widths)

                // -----------------------------------------------------------------
                // 2. Copy columns (including column widths and cell data/formats)
                //    Destination columns start at index 0.
                // -----------------------------------------------------------------
                destSheet.Cells.CopyColumns(
                    srcSheet.Cells,          // source cells
                    srcFirstColumn,         // source start column
                    0,                      // destination start column
                    columnCount);           // number of columns to copy

                // Save the destination workbook
                destWorkbook.Save(destinationPath, SaveFormat.Xlsx);
                Console.WriteLine($"Destination workbook saved to {destinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
