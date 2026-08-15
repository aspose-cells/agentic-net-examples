// Title: Copy a Range to a New Workbook and Apply a Table Style with Aspose.Cells for .NET
// Description: Demonstrates how to copy a defined range (A1:D5) from one workbook to another using Aspose.Cells, convert the copied range into a ListObject, apply a built‑in table style (TableStyleMedium2) that enables filtering and sorting, and save the result as an Excel file.
// Keywords: Aspose.Cells copy range | Aspose.Cells ListObject | Aspose.Cells table style | PasteOptions KeepOldTables | C# Excel range copy | apply table style .NET | Excel filtering sorting Aspose | convert range to table Aspose.Cells
// Common Searches: How to copy a range to another workbook with Aspose.Cells | Aspose.Cells create ListObject from copied range | Apply built‑in table style after copying range in C# | Enable filter and sort on copied Excel range using Aspose | PasteOptions KeepOldTables example Aspose.Cells
// Developer Intent: Copy a specific cell range to a new workbook and turn it into a styled table that supports filtering and sorting.
// Use Cases: Generate a report by copying data from a template sheet into a fresh workbook and automatically adding a sortable, filterable table. | Migrate a data block to a new file while preserving formatting and enabling end‑user analysis through table features. | Create a reusable utility that copies any range to another workbook and applies a consistent medium‑style table for UI uniformity.
// AI Prompts: Provide a reusable C# method that accepts a source workbook, source range address, and destination path, then copies the range and adds a table with a specified style. | Show how to define and apply a custom table style after copying a range instead of using a built‑in style. | Explain how to copy multiple non‑contiguous ranges from one workbook and create separate ListObjects for each in the destination workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyWithTableStyle
{
    // Demonstrates how to copy a defined range (A1:D5) from one workbook to another using Aspose.Cells, convert the copied range into a ListObject, apply a built‑in table style (TableStyleMedium2) that enables filtering and sorting, and save the result as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook and fill sample data ----------
                Workbook sourceWb = new Workbook();                         // create source workbook
                Worksheet sourceWs = sourceWb.Worksheets[0];

                // Fill a 5x4 block with sample values (A1:D5)
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        sourceWs.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define the source range to be copied
                AsposeRange sourceRange = sourceWs.Cells.CreateRange("A1:D5");

                // ---------- Create destination workbook ----------
                Workbook destWb = new Workbook();                           // create destination workbook
                Worksheet destWs = destWb.Worksheets[0];

                // Define the destination range (same size, starting at A1)
                AsposeRange destRange = destWs.Cells.CreateRange("A1:D5");

                // ---------- Copy source range to destination range ----------
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All,      // copy everything (values, formulas, formats, etc.)
                    KeepOldTables = true            // preserve any existing tables in the destination
                };
                destRange.Copy(sourceRange, pasteOptions); // copy with options

                // ---------- Convert the copied range into a table (ListObject) ----------
                // Add a ListObject that covers the destination range
                int tableIndex = destWs.ListObjects.Add(
                    destRange.FirstRow,
                    destRange.FirstColumn,
                    destRange.RowCount,
                    destRange.ColumnCount,
                    true); // true => has header row

                ListObject table = destWs.ListObjects[tableIndex];

                // Apply a built‑in table style (enables filtering and sorting)
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Ensure the style is applied to the underlying range
                table.ApplyStyleToRange(); // apply style to the range

                // ---------- Save the destination workbook ----------
                string outputPath = "CopiedRangeWithTableStyle.xlsx";
                destWb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
