using System;
using Aspose.Cells;

namespace AutoFitRowHeightValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate cells with long wrapped text to trigger auto‑fit
            cells["A1"].PutValue("This is a very long piece of text that should cause the row to expand significantly when auto‑fitted.");
            cells["B1"].PutValue("Another long text that will be wrapped across multiple lines to test row height limits.");
            cells["A2"].PutValue("Short text");
            cells["B2"].PutValue("More short text");

            // Enable text wrapping for the cells
            Style wrapStyle = cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            cells["A1"].SetStyle(wrapStyle);
            cells["B1"].SetStyle(wrapStyle);

            // Define the maximum allowed row height (in points)
            const double maxAllowedHeight = 40.0; // Example limit

            // Configure AutoFitterOptions with MaxRowHeight (rule)
            AutoFitterOptions options = new AutoFitterOptions
            {
                MaxRowHeight = maxAllowedHeight,
                OnlyAuto = true // Fit only rows that are not custom‑height
            };

            // Auto‑fit rows using the options (rule)
            worksheet.AutoFitRows(options);

            // Validate that no row exceeds the maximum height
            int lastRow = cells.MaxDataRow; // Highest row that contains data
            for (int rowIndex = 0; rowIndex <= lastRow; rowIndex++)
            {
                double rowHeight = cells.GetRowHeight(rowIndex); // Height in points
                if (rowHeight > maxAllowedHeight)
                {
                    Console.WriteLine($"Row {rowIndex + 1} height ({rowHeight:F2} pt) exceeds the limit of {maxAllowedHeight} pt.");
                }
                else
                {
                    Console.WriteLine($"Row {rowIndex + 1} height is within limit: {rowHeight:F2} pt.");
                }
            }

            // Save the workbook (save rule)
            workbook.Save("AutoFitRowHeightValidated.xlsx");
        }
    }
}