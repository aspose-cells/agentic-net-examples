// Title: C# – Convert Excel Table to Range, Keep First 5 Rows Formatting, Save as ODS with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a ListObject (table), apply a built‑in style, convert the table to a normal range, clear formatting from rows 6‑10 while preserving the first five rows, and export the result as an ODS file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# convert table to range | preserve first rows formatting Aspose | clear formats rows Aspose.Cells | save workbook as ODS .NET | ListObject ConvertToRange ODS export | TableStyleMedium2 Aspose.Cells
// Common Searches: Aspose.Cells convert ListObject to range keep header formatting | C# clear table formatting after ConvertToRange | Export Excel table as ODS with Aspose.Cells | How to keep first N rows styled when converting table to range | Aspose.Cells OdsSaveOptions example
// Developer Intent: Convert a ListObject to a regular cell range, retain the table style for the first five rows, remove styling from the remaining rows, and save the workbook as an ODS document.
// Use Cases: Generate a styled preview where only the header and top five data rows keep the table appearance before sharing the file with LibreOffice. | Create a template that uses a table style for visual consistency, then strip formatting from rows beyond the preview section for downstream ODS processing. | Export Excel data to ODS while preserving essential formatting for reporting headers and a limited data slice.
// AI Prompts: Write C# code using Aspose.Cells to convert a ListObject to a range, keep formatting for the first N rows, and save as ODS. | Show how to clear formatting from rows 6‑10 after converting an Excel table to a range with Aspose.Cells. | Explain OdsSaveOptions settings for exporting a workbook that originally contained a styled table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace AsposeCellsTableToRange
{
    // Demonstrates how to create a workbook, add a ListObject (table), apply a built‑in style, convert the table to a normal range, clear formatting from rows 6‑10 while preserving the first five rows, and export the result as an ODS file using Aspose.Cells for .NET.
    public class ConvertTableToRangeWithFormatting
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (10 rows, 3 columns)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            for (int row = 2; row <= 10; row++)
            {
                sheet.Cells[row - 1, 0].PutValue(row - 1);                     // ID
                sheet.Cells[row - 1, 1].PutValue($"Person {row - 1}");        // Name
                sheet.Cells[row - 1, 2].PutValue((row - 1) * 10);             // Score
            }

            // Create a table (ListObject) that covers the whole data range
            int tableIndex = sheet.ListObjects.Add(0, 0, 9, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply a built‑in table style (affects all rows initially)
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Convert the table to a normal range (keeps formatting for all rows)
            table.ConvertToRange();

            // Clear formatting for rows beyond the first five (rows 6‑10, zero‑based index 5‑9)
            // Parameters: startRow, startColumn, totalRows, totalColumns
            sheet.Cells.ClearFormats(5, 0, 5, 3); // 5 rows, 3 columns

            // Save the workbook as ODS using default OdsSaveOptions
            OdsSaveOptions odsOptions = new OdsSaveOptions();
            workbook.Save("TableConvertedRange.ods", odsOptions);
        }
    }
}
