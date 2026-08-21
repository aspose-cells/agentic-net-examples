// Title: Validate Auto‑Fitted Row Height Does Not Exceed Max Limit with Aspose.Cells for .NET
// Description: Creates a workbook, writes long wrapped text to cells A1 and A2, enables text wrapping, sets a maximum row height via AutoFitterOptions.MaxRowHeight, auto‑fits rows 0‑1, reads each row's height with Cells.GetRowHeight, compares it to the limit, logs the outcome, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | AutoFitRows | MaxRowHeight | row height validation | limit row height | retrieve row height | Excel row height constraint | auto‑fit rows example
// Common Searches: Aspose.Cells limit row height after AutoFitRows | C# check row height after auto‑fit | Set maximum row height with AutoFitterOptions | GetRowHeight after AutoFitRows Aspose.Cells | Validate row height constraint in .NET Excel
// Developer Intent: Confirm that rows auto‑fitted by Aspose.Cells stay within a predefined maximum height.
// Use Cases: Apply text wrapping, auto‑fit rows, and enforce a height ceiling (e.g., 40 points) using AutoFitterOptions. | Iterate over a range of rows to detect and log any that exceed the allowed height. | Save the workbook after validation to preserve the enforced row‑height limits.
// AI Prompts: Generate C# code that auto‑fits rows with a MaxRowHeight of 30 points and throws an exception if any row exceeds the limit. | Write a method that accepts a worksheet and a maximum height, auto‑fits a specified row range, and returns a list of rows violating the constraint. | Explain how AutoFitterOptions.MaxRowHeight works with text wrapping and how to retrieve actual row heights after auto‑fit.

using System;
using Aspose.Cells;

namespace AutoFitRowHeightValidation
{
    // Creates a workbook, writes long wrapped text to cells A1 and A2, enables text wrapping, sets a maximum row height via AutoFitterOptions.MaxRowHeight, auto‑fits rows 0‑1, reads each row's height with Cells.GetRowHeight, compares it to the limit, logs the outcome, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with long wrapped text to trigger auto‑fit
            cells["A1"].PutValue("This is a very long text that should cause the row height to increase significantly when auto‑fitted.");
            cells["A2"].PutValue("Another long text with line breaks\nto further test the auto‑fit behavior.");
            // Enable text wrapping so that height changes
            Style wrapStyle = cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            cells["A1"].SetStyle(wrapStyle);
            cells["A2"].SetStyle(wrapStyle);

            // Define the maximum allowed row height (in points)
            double maxAllowedHeight = 40.0; // Example limit

            // Configure AutoFitterOptions with MaxRowHeight
            AutoFitterOptions options = new AutoFitterOptions
            {
                MaxRowHeight = maxAllowedHeight,
                OnlyAuto = true
            };

            // Auto‑fit rows 0 to 1 using the options (auto‑fit rule)
            sheet.AutoFitRows(0, 1, options);

            // Validate that each row height does not exceed the maximum
            for (int row = 0; row <= 1; row++)
            {
                double actualHeight = cells.GetRowHeight(row); // get row height rule
                if (actualHeight > maxAllowedHeight)
                {
                    Console.WriteLine($"Row {row} height {actualHeight:F2} exceeds the limit of {maxAllowedHeight} points.");
                }
                else
                {
                    Console.WriteLine($"Row {row} height {actualHeight:F2} is within the allowed limit.");
                }
            }

            // Save the workbook (save rule)
            workbook.Save("AutoFitRowHeightValidation.xlsx");
        }
    }
}
