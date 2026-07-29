// Title: C# – Limit Row Height When Auto‑Fitting Rows with Aspose.Cells AutoFitterOptions
// Description: Shows how to use AutoFitterOptions (MaxRowHeight and OnlyAuto) to auto‑fit rows in a worksheet while preventing rows from exceeding a defined height, and keeping manually sized rows unchanged.
// Keywords: Aspose.Cells | C# | AutoFitterOptions | MaxRowHeight | OnlyAuto | auto fit rows | limit row height | Excel row height | wrap text | prevent tall rows
// Common Searches: Aspose.Cells limit row height auto fit | C# AutoFitterOptions MaxRowHeight example | prevent excessively tall rows Aspose.Cells | OnlyAuto property usage in AutoFitRows | set maximum row height before auto‑fit Aspose
// Developer Intent: Cap the height of rows during AutoFitRows so that wrapped text does not create overly tall rows.
// Use Cases: Apply a 50‑point maximum height to rows that contain wrapped text. | Auto‑fit rows only when they have not been given a custom height. | Generate Excel reports with long descriptions that stay within a consistent row height.
// AI Prompts: How do I use Aspose.Cells AutoFitterOptions to set MaxRowHeight and OnlyAuto in C#? | Provide a C# code snippet that limits row height to 40 points when auto‑fitting wrapped text with Aspose.Cells. | Explain the effect of the OnlyAuto property on rows that already have a custom height.

using System;
using Aspose.Cells;

namespace AsposeCellsMaxRowHeightDemo
{
    // Shows how to use AutoFitterOptions (MaxRowHeight and OnlyAuto) to auto‑fit rows in a worksheet while preventing rows from exceeding a defined height, and keeping manually sized rows unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a long text that would normally cause a very tall row when wrapped
            worksheet.Cells["A1"].PutValue("This is a very long piece of text that will be wrapped and could cause the row to become excessively tall if auto‑fitted without limits.");

            // Enable text wrapping for the cell
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Optionally set an initial small row height
            worksheet.Cells.SetRowHeight(0, 10); // 10 points

            // Create AutoFitterOptions and set the maximum row height (in points)
            AutoFitterOptions options = new AutoFitterOptions
            {
                MaxRowHeight = 50,   // Limit rows to a maximum of 50 points
                OnlyAuto = true      // Apply only to rows that have not been custom‑sized
            };

            // Auto‑fit all rows using the options with the max height restriction
            worksheet.AutoFitRows(options);

            // Save the workbook
            workbook.Save("MaxRowHeightDemo.xlsx");
        }
    }
}
