using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFormattingAndNamedRanges
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in a 5x4 area (A1:D5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a named range "DataRange" that refers to A1:D5
            AsposeRange dataRange = cells.CreateRange("A1", "D5");
            dataRange.Name = "DataRange";

            // Apply a style (yellow background, bold font) to the entire named range
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.Yellow;
            style.Font.IsBold = true;
            dataRange.SetStyle(style);

            // Clear formatting from a sub‑range B2:C3 within the named range
            cells.ClearFormats(1, 1, 2, 2); // rows and columns are zero‑based

            // Extend the named range to include two more columns and one more row (A1:E6)
            Name namedRange = workbook.Worksheets.Names["DataRange"];
            string newRef = $"={sheet.Name}!$A$1:$E$6";
            namedRange.RefersTo = newRef;

            // Verify the updated range by retrieving it
            AsposeRange updatedRange = namedRange.GetRange();
            Console.WriteLine($"Updated named range address: {updatedRange.RefersTo}");

            // Save the workbook
            workbook.Save("FormattedNamedRanges.xlsx");
        }
    }
}