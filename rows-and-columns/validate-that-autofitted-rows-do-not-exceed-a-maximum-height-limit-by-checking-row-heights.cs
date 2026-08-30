// Title: C# – Validate that auto‑fitted rows do not exceed a specified maximum height using Aspose.Cells AutoFitterOptions
// AI Prompts: Generate C# code that applies AutoFitterOptions with MaxRowHeight, auto‑fits rows, and then iterates through each data row to compare its height against the limit. | Show how to log rows whose actual height is greater than a defined point value after calling Worksheet.AutoFitRows with a height cap. | Provide an example that wraps cell text, sets a maximum row height, performs auto‑fit, and outputs validation results for each row.
// Common Searches: Aspose.Cells C# enforce maximum row height when using AutoFitRows | how to check row height after AutoFitterOptions AutoFitRows in .NET | C# code to detect rows taller than a given point size in an Aspose.Cells workbook | retrieve row height for each data row after auto‑fitting in Aspose.Cells | validate that auto‑fitted rows stay within 40 points using Aspose.Cells
// Tags: apply MaxRowHeight constraint Aspose.Cells | retrieve row height after AutoFitRows C# | enforce row height cap Aspose.Cells .NET | validate row height limit using AutoFitterOptions | wrap text and limit row height Aspose.Cells

using System;
using Aspose.Cells;

namespace AutoFitRowHeightValidation
{
    // The example creates a workbook, writes long wrapped text, configures AutoFitterOptions with a MaxRowHeight, auto‑fits rows, iterates through all data rows to compare actual heights with the defined limit, logs whether each row is within or exceeds the cap, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate cells with long wrapped text to trigger row height increase
            cells["A1"].PutValue("This is a very long piece of text that should cause the row to expand when auto‑fitted. " +
                                 "It contains multiple sentences to ensure the height grows significantly.");
            cells["B1"].PutValue("Another long text in the adjacent column to further influence the row height.");
            // Enable text wrapping for the cells
            Style wrapStyle = cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            cells["A1"].SetStyle(wrapStyle);
            cells["B1"].SetStyle(wrapStyle);

            // Define the maximum allowed row height (in points)
            double maxAllowedHeight = 40.0; // Example: 40 points

            // Configure AutoFitterOptions with the MaxRowHeight limit
            AutoFitterOptions options = new AutoFitterOptions
            {
                MaxRowHeight = maxAllowedHeight,
                OnlyAuto = true // Fit only rows that are not manually sized
            };

            // Auto‑fit rows using the options (rule: AutoFitRows(AutoFitterOptions))
            worksheet.AutoFitRows(options);

            // Validate that no row exceeds the specified maximum height
            int lastRow = cells.MaxDataRow; // Check rows that contain data
            for (int rowIndex = 0; rowIndex <= lastRow; rowIndex++)
            {
                double actualHeight = cells.GetRowHeight(rowIndex); // Rule: GetRowHeight(int)
                if (actualHeight > maxAllowedHeight)
                {
                    Console.WriteLine($"Row {rowIndex} height {actualHeight:F2} exceeds the limit of {maxAllowedHeight} points.");
                }
                else
                {
                    Console.WriteLine($"Row {rowIndex} height {actualHeight:F2} is within the allowed limit.");
                }
            }

            // Save the workbook (rule: workbook.Save)
            workbook.Save("AutoFitRowHeightValidated.xlsx");
        }
    }
}
