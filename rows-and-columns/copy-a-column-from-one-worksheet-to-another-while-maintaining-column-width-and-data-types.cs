// Title: Copy a worksheet column while preserving its width and data types using Aspose.Cells for .NET (C#)
// AI Prompts: Copy column A from a source worksheet to a destination worksheet, keeping the original column width and all data types, with Aspose.Cells in C#. | Use Cells.CopyColumns together with PasteOptions.All to duplicate a column—including values, formats, and column width—between two workbooks. | Transfer a column between workbooks while preserving string, numeric, date, and boolean types using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# copy column preserving column width | how to keep data types when copying columns between worksheets Aspose.Cells | Cells.CopyColumns example with formatting and width Aspose.Cells | PasteOptions.All usage for column copy in Aspose.Cells .NET | transfer column from one workbook to another preserving styles Aspose.Cells
// Tags: copy column with column width Aspose.Cells | preserve data types Cells.CopyColumns | PasteOptions.All column copy C# | transfer column between workbooks Aspose.Cells | set column width characters Aspose.Cells | duplicate worksheet column formatting Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsColumnCopyDemo
{
    // The example creates a source workbook, fills column A with various data types, sets its width to 25 characters, and then copies that column to a new workbook using Cells.CopyColumns with PasteOptions.All, ensuring values, formats, and column width are retained before saving both files.
    class Program
    {
        static void Main()
        {
            // Create source workbook and get its first worksheet
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Populate source column (A) with different data types
            srcSheet.Cells["A1"].PutValue("Text");          // string
            srcSheet.Cells["A2"].PutValue(12345);           // integer
            srcSheet.Cells["A3"].PutValue(3.14159);         // double
            srcSheet.Cells["A4"].PutValue(DateTime.Now);    // DateTime
            srcSheet.Cells["A5"].PutValue(true);            // boolean

            // Set column width for column A (index 0) in characters
            srcSheet.Cells.SetColumnWidth(0, 25);

            // Create destination workbook and get its first worksheet
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Prepare paste options to copy all data and formats (including column width)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All   // copies values, formulas, formats, column widths, etc.
            };

            // Copy the first column (index 0) from source to destination
            // Parameters: sourceCells, sourceColumnIndex, destinationColumnIndex, columnNumber, pasteOptions
            destSheet.Cells.CopyColumns(
                srcSheet.Cells,
                0,          // source column index (A)
                0,          // destination column index (A)
                1,          // number of columns to copy
                pasteOptions);

            // Save both workbooks to verify the result
            srcWorkbook.Save("SourceWorkbook.xlsx");
            destWorkbook.Save("DestinationWorkbook.xlsx");

            Console.WriteLine("Column copied successfully with data types and column width preserved.");
        }
    }
}
