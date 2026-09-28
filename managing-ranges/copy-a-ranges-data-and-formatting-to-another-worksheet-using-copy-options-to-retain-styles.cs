// Title: Copy a cell range with data and formatting to another worksheet using Aspose.Cells PasteOptions in C#
// AI Prompts: Copy the range A1:B3 from the Source sheet to cell C5 on the Destination sheet while preserving fonts, colors, and background using Aspose.Cells PasteOptions. | Configure a PasteOptions object to retain all cell styles, column widths, and row heights when moving a range between worksheets in a .NET workbook. | Demonstrate how to create a destination range, perform a full‑style copy, and save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy range to another sheet keeping cell styles | How to preserve column width and row height when copying cells with Aspose.Cells | Using PasteOptions.PasteType.All to duplicate formatted range in .NET workbook | Copy styled range from one worksheet to another using Aspose.Cells API
// Tags: copy range with formatting Aspose.Cells C# | PasteOptions PasteType.All usage | retain column width Aspose.Cells | move styled cells between worksheets | export workbook to xlsx Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, adds a source and a destination worksheet, fills A1:B3 with values, applies bold blue font on a yellow background, then copies that range to C5 on the destination sheet using PasteOptions configured with PasteType.All to keep all data, styles, column widths, and row heights, and finally saves the file as CopyRangeWithFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add source and destination worksheets
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";
            Worksheet destinationSheet = workbook.Worksheets.Add("Destination");

            // Populate source range with data
            sourceSheet.Cells["A1"].PutValue("Header");
            sourceSheet.Cells["A2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue(20);
            sourceSheet.Cells["B2"].PutValue(30);
            sourceSheet.Cells["B3"].PutValue(40);

            // Apply formatting to the source range
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            style.Font.Color = Color.Blue;
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            sourceSheet.Cells.CreateRange("A1:B3").ApplyStyle(style, new StyleFlag() { All = true });

            // Define the source range (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range srcRange = sourceSheet.Cells.CreateRange("A1:B3");

            // Configure paste options (retain all data and formatting)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All
                // Column width and row height are copied by default
            };

            // Create a destination range starting at C5 (row index 4, column index 2)
            Aspose.Cells.Range destRange = destinationSheet.Cells.CreateRange(4, 2, srcRange.RowCount, srcRange.ColumnCount);

            // Perform the copy
            srcRange.Copy(destRange, pasteOptions);

            // Save the workbook
            workbook.Save("CopyRangeWithFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
