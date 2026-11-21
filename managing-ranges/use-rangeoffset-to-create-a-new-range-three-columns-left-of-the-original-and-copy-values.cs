// Title: Copy only cell values to a range three columns left with Range offset and PasteOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy values to a range offset three columns left | Create a left‑shifted range and copy only values using PasteOptions in Aspose.Cells | How to use Range.CreateRange with column offset in Aspose.Cells .NET | Copy values from D5:F7 to A5:C7 with Aspose.Cells C# example | Offset a range by columns and paste values only in Aspose.Cells
// Tags: range offset copy Aspose.Cells | CreateRange column shift C# | PasteOptions values only Aspose.Cells | copy values between Excel ranges Aspose.Cells | leftward range offset Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills the range D5:F7 with sample data, defines an offset range three columns to the left (A5:C7) by adjusting the first column index, and copies only the cell values from the original range to the offset range using PasteOptions with PasteType.Values. The workbook is then saved as OffsetCopy.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the original range D5:F7 with sample data
            // (Rows and columns are zero‑based in the API)
            for (int row = 4; row <= 6; row++)          // rows 5‑7
            {
                for (int col = 3; col <= 5; col++)      // columns D‑F
                {
                    sheet.Cells[row, col].PutValue($"R{row - 3}C{col - 2}");
                }
            }

            // Get the original range D5:F7
            Aspose.Cells.Range originalRange = sheet.Cells.CreateRange("D5", "F7");

            // Define the offset range three columns to the left (A5:C7)
            int offsetRow = originalRange.FirstRow;
            int offsetColumn = originalRange.FirstColumn - 3; // shift left by 3 columns
            int rowCount = originalRange.RowCount;
            int columnCount = originalRange.ColumnCount;
            Aspose.Cells.Range offsetRange = sheet.Cells.CreateRange(offsetRow, offsetColumn, rowCount, columnCount);

            // Copy only the values from the original range to the offset range
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Values
            };
            offsetRange.Copy(originalRange, pasteOptions);

            // Save the workbook
            workbook.Save("OffsetCopy.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
