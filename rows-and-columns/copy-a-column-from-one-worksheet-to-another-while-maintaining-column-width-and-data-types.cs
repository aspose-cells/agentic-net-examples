// Title: Copy Column Between Worksheets in Aspose.Cells for .NET – Preserve Data Types & Column Width
// Description: Demonstrates how to use Aspose.Cells' CopyColumns method with PasteOptions.PasteType.All to transfer a column from one worksheet to another while keeping original data types, formatting, and column width, then saves both workbooks.
// Keywords: Aspose.Cells | C# | CopyColumns | PasteOptions | column width | preserve data types | worksheet copy | Excel automation | Aspose.Cells .NET example | copy column formatting
// Common Searches: Aspose.Cells copy column preserve formatting | CopyColumns method C# example | how to keep column width when copying in Aspose.Cells | copy worksheet column with data types Aspose.Cells | PasteOptions.PasteType.All usage
// Developer Intent: Transfer a column from a source worksheet to a target worksheet without losing its data types, formatting, or column width.
// Use Cases: Reuse a formatted data column from a template workbook in multiple report workbooks. | Duplicate configuration columns across several Excel files while maintaining custom widths and styles. | Create a new workbook that mirrors the layout of an existing sheet for consistent data presentation.
// AI Prompts: Write C# code that copies several adjacent columns with all formatting and column widths using Aspose.Cells. | Explain the effect of PasteOptions.PasteType.All on the CopyColumns operation in Aspose.Cells. | Provide a step‑by‑step guide to copy a column and then change its date format to "yyyy‑MM‑dd" using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsColumnCopyDemo
{
    // Demonstrates how to use Aspose.Cells' CopyColumns method with PasteOptions.PasteType.All to transfer a column from one worksheet to another while keeping original data types, formatting, and column width, then saves both workbooks.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook srcWorkbook = new Workbook();                     // create
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Populate source column (A) with different data types
            srcSheet.Cells["A1"].PutValue("Header");                 // string
            srcSheet.Cells["A2"].PutValue(12345);                    // integer
            srcSheet.Cells["A3"].PutValue(123.456);                  // double
            srcSheet.Cells["A4"].PutValue(DateTime.Now);             // DateTime
            srcSheet.Cells["A5"].PutValue(true);                     // boolean

            // Set a custom column width (in characters) for column A
            srcSheet.Cells.SetColumnWidth(0, 25); // column index 0 = A

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();                    // create
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Prepare paste options to copy everything (data, formats, column width)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All   // copies data, formats, and column widths
            };

            // Copy the first column (index 0) from source to destination column index 2 (C)
            // Copy 1 column, preserving data types and column width
            destSheet.Cells.CopyColumns(
                srcSheet.Cells,   // source cells
                0,                // source column index (A)
                2,                // destination column index (C)
                1,                // number of columns to copy
                pasteOptions);   // paste options

            // ---------- Save workbooks ----------
            srcWorkbook.Save("SourceWorkbook.xlsx");   // save source
            destWorkbook.Save("DestinationWorkbook.xlsx"); // save destination

            Console.WriteLine("Column copied successfully with data types and column width preserved.");
        }
    }
}
