// Title: Auto‑fit columns B‑D and then enforce a 120‑pixel width for each column using Aspose.Cells for .NET (C#)
// AI Prompts: Auto‑fit columns B, C, and D in a worksheet and then set each column to exactly 120 pixels using Aspose.Cells in C#. | After calling AutoFitColumn for a range, apply SetColumnWidthPixel to the same columns to achieve precise pixel alignment in a .NET workbook.
// Common Searches: Aspose.Cells C# set exact pixel width after AutoFitColumn for multiple columns | How to force column width to 120 pixels after auto‑fitting in Aspose.Cells .NET | Batch adjust column widths to a specific pixel size after autofit using Aspose.Cells | SetColumnWidthPixel on a range of columns following AutoFitColumn in C# | Precise column sizing in Excel with Aspose.Cells after auto‑fit
// Tags: auto-fit column width Aspose.Cells C# | set column width pixel Aspose.Cells | multiple column pixel sizing .NET | column width alignment after autofit | fixed pixel column width Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, populates columns B‑D with data, auto‑fits those columns, then overrides the auto‑fit result by setting each column to a fixed 120‑pixel width using SetColumnWidthPixel, and finally saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Sample data in columns B to D (indexes 1‑3)
            sheet.Cells["B1"].PutValue("Header1");
            sheet.Cells["C1"].PutValue("Header2");
            sheet.Cells["D1"].PutValue("Header3");
            sheet.Cells["B2"].PutValue("Longer text in column B");
            sheet.Cells["C2"].PutValue(12345);
            sheet.Cells["D2"].PutValue(DateTime.Now);

            // Auto‑fit columns B‑D (zero‑based column indexes 1‑3)
            for (int col = 1; col <= 3; col++)
            {
                sheet.AutoFitColumn(col);
            }

            // Set precise pixel width after auto‑fit
            int targetPixelWidth = 120; // desired width in pixels
            for (int col = 1; col <= 3; col++)
            {
                // Use Cells collection to set column width in pixels
                sheet.Cells.SetColumnWidthPixel(col, targetPixelWidth);
            }

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
