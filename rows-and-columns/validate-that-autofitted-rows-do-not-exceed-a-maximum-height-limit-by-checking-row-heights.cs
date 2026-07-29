// Title: C# – Validate Max Row Height After AutoFitRows with Aspose.Cells
// Description: Shows how to configure AutoFitterOptions with a MaxRowHeight, auto‑fit rows that contain wrapped text, and programmatically verify that each row stays within the defined height limit using Aspose.Cells for .NET.
// Keywords: Aspose.Cells AutoFitRows C# | MaxRowHeight | AutoFitterOptions | row height validation | limit row height .NET | check row height after autofit | C# spreadsheet row height limit
// Common Searches: Aspose.Cells limit row height | C# AutoFitRows maximum height | How to enforce max row height in Aspose.Cells | Validate row height after autofit C# | AutoFitterOptions MaxRowHeight example
// Developer Intent: The developer needs to ensure that rows auto‑fitted by Aspose.Cells do not exceed a predefined maximum height.
// Use Cases: Apply a maximum row height while auto‑fitting rows that contain wrapped text to keep the layout consistent. | Iterate through rows after AutoFitRows to log or handle rows whose height exceeds the allowed limit. | Save the workbook only after confirming all rows meet the height constraint.
// AI Prompts: Generate C# code that uses Aspose.Cells to auto‑fit rows with a maximum height of 30 points and reports any rows that exceed this limit. | Provide an example of configuring AutoFitterOptions to apply MaxRowHeight only to rows that are not manually sized. | Show how to retrieve row heights after AutoFitRows and compare them to a threshold in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AutoFitRowHeightValidation
{
    // Shows how to configure AutoFitterOptions with a MaxRowHeight, auto‑fit rows that contain wrapped text, and programmatically verify that each row stays within the defined height limit using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some rows with long wrapped text to trigger auto‑fit
            for (int i = 0; i < 5; i++)
            {
                string longText = "This is a very long piece of text that should cause the row height to increase significantly. Row " + i;
                cells[i, 0].PutValue(longText);
                Style style = cells[i, 0].GetStyle();
                style.IsTextWrapped = true;          // Enable wrapping so height can grow
                cells[i, 0].SetStyle(style);
            }

            // Define the maximum allowed row height (in points)
            double maxAllowedHeight = 40.0; // Example limit

            // Configure AutoFitterOptions with the MaxRowHeight limit
            AutoFitterOptions options = new AutoFitterOptions
            {
                MaxRowHeight = maxAllowedHeight,
                OnlyAuto = true                     // Fit only rows that are not manually sized
            };

            // Auto‑fit all rows using the options (the MaxRowHeight will be enforced)
            worksheet.AutoFitRows(options);

            // Validate that no row exceeds the maximum height
            int maxRow = cells.MaxDataRow; // Highest row that contains data
            for (int row = 0; row <= maxRow; row++)
            {
                double actualHeight = cells.GetRowHeight(row); // Height in points
                if (actualHeight > maxAllowedHeight)
                {
                    Console.WriteLine($"Row {row} height {actualHeight:F2} exceeds the limit of {maxAllowedHeight} points.");
                }
                else
                {
                    Console.WriteLine($"Row {row} height {actualHeight:F2} is within the allowed limit.");
                }
            }

            // Save the workbook (output path can be adjusted as needed)
            workbook.Save("AutoFitRowsValidated.xlsx");
        }
    }
}
