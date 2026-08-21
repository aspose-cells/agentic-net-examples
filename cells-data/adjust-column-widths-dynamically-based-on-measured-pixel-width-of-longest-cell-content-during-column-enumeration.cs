// Title: C# – Auto‑Fit Excel Columns by Pixel Width Using Aspose.Cells
// Description: Creates a workbook, fills a 10‑row × 5‑column range with varied‑length strings, auto‑fits each column for the defined rows, reads the resulting pixel width, adds optional padding, and sets the final column width before saving the file.
// Keywords: Aspose.Cells C# column width pixel | AutoFitColumn pixel measurement | set column width Aspose.Cells | dynamic Excel column sizing | GetColumnWidthPixel | SetColumnWidthPixel | Excel column padding C#
// Common Searches: how to set Excel column width in pixels with Aspose.Cells | auto‑fit column and add extra pixels Aspose.Cells .NET | retrieve column pixel width after AutoFit Aspose | adjust column widths based on longest cell content Aspose.Cells
// Developer Intent: Programmatically size each column to the pixel width of its longest cell content, optionally adding a small padding.
// Use Cases: Generate reports where columns automatically expand to display the longest text without truncation. | Import datasets of unknown length and ensure consistent layout by adjusting column widths on the fly. | Apply a uniform pixel buffer after AutoFit to improve readability in the final Excel file.
// AI Prompts: Write C# code that uses Aspose.Cells to AutoFit a column range, obtain the pixel width, add 5 px padding, and set the new width. | Explain the interaction between GetColumnWidthPixel and SetColumnWidthPixel after calling AutoFitColumn in Aspose.Cells. | Suggest a method to calculate padding based on font size when adjusting column widths with Aspose.Cells.

using System;
using Aspose.Cells;

namespace DynamicColumnWidthDemo
{
    // Creates a workbook, fills a 10‑row × 5‑column range with varied‑length strings, auto‑fits each column for the defined rows, reads the resulting pixel width, adds optional padding, and sets the final column width before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: populate rows and columns with varying length strings
            int firstRow = 0;
            int lastRow = 9;   // 10 rows (0‑based)
            int firstCol = 0;
            int lastCol = 4;   // 5 columns (0‑based)

            string[] sampleTexts = new string[]
            {
                "Short",
                "Medium length text",
                "A considerably longer piece of text that should expand the column width",
                "Tiny",
                "Extremely long text that will definitely require a wider column to be fully visible in the sheet"
            };

            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {
                    // Cycle through sample texts to create varied content
                    string text = sampleTexts[(row + col) % sampleTexts.Length];
                    cells[row, col].PutValue(text);
                }
            }

            // Iterate through each column, auto‑fit based on its content,
            // then retrieve the calculated pixel width and set it explicitly
            for (int col = firstCol; col <= lastCol; col++)
            {
                // Auto‑fit the column for the defined row range
                sheet.AutoFitColumn(col, firstRow, lastRow);

                // Get the width that AutoFit calculated (in pixels)
                int pixelWidth = sheet.Cells.GetColumnWidthPixel(col);

                // Optional: add a small padding (e.g., 5 pixels) for visual comfort
                int paddedWidth = pixelWidth + 5;

                // Apply the pixel width to the column
                sheet.Cells.SetColumnWidthPixel(col, paddedWidth);
            }

            // Save the workbook
            workbook.Save("DynamicColumnWidth.xlsx");
        }
    }
}
