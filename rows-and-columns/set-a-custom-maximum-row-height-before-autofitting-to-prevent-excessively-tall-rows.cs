// Title: C# – Limit Row Height with AutoFitterOptions MaxRowHeight in Aspose.Cells
// Description: Demonstrates how to set a maximum row height before calling AutoFitRows, using AutoFitterOptions (MaxRowHeight = 40, OnlyAuto = true) to keep wrapped text from creating overly tall rows, then saves the workbook.
// Keywords: Aspose.Cells AutoFitRows | AutoFitterOptions MaxRowHeight | C# limit row height | prevent tall rows Excel | OnlyAuto property | wrap text row height | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set maximum row height auto fit | AutoFitterOptions MaxRowHeight C# example | prevent rows from expanding too much Aspose.Cells | OnlyAuto option usage in AutoFitRows | limit row height when wrapping text in Excel using Aspose
// Developer Intent: Apply a custom maximum height to rows before auto‑fitting to avoid excessive row expansion.
// Use Cases: Keep rows with long wrapped text under a defined height (e.g., 40 points) for consistent report layouts. | Auto‑fit only rows that have not been manually sized, preserving user‑defined heights. | Generate Excel files where visual height constraints are required for printing or PDF conversion.
// AI Prompts: Show C# code that uses Aspose.Cells AutoFitterOptions to cap row height at 30 points and auto‑fit only rows without a preset height. | Write a .NET example that wraps text in a cell, limits the row height, and saves the workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to set a maximum row height before calling AutoFitRows, using AutoFitterOptions (MaxRowHeight = 40, OnlyAuto = true) to keep wrapped text from creating overly tall rows, then saves the workbook.
    public class MaxRowHeightAutoFitDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add long wrapped text to demonstrate row auto‑fitting
                worksheet.Cells["A1"].PutValue(
                    "This is a very long text that would normally cause the row to become excessively tall after auto‑fitting.");
                Style style = worksheet.Cells["A1"].GetStyle();
                style.IsTextWrapped = true;
                worksheet.Cells["A1"].SetStyle(style);

                // Optionally set an initial small row height
                worksheet.Cells.SetRowHeight(0, 10);

                // Create AutoFitterOptions with a maximum row height limit
                AutoFitterOptions options = new AutoFitterOptions
                {
                    MaxRowHeight = 40, // limit maximum row height to 40 points
                    OnlyAuto = true    // apply only to rows without custom height
                };

                // Auto‑fit rows using the specified options
                worksheet.AutoFitRows(options);

                // Display the resulting row height
                Console.WriteLine(
                    "Row 0 height after auto‑fit with MaxRowHeight=40: " +
                    worksheet.Cells.GetRowHeight(0));

                // Save the workbook
                workbook.Save("MaxRowHeightAutoFitDemo.xlsx");
                Console.WriteLine("Workbook saved as MaxRowHeightAutoFitDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MaxRowHeightAutoFitDemo.Run();
        }
    }
}
