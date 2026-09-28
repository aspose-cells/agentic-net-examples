// Title: How to compare original and adjusted row heights after using AutoFitRows with Aspose.Cells in C#
// AI Prompts: Write C# code that reads a row's height, calls worksheet.AutoFitRows(), then outputs the before and after heights using Aspose.Cells. | Show an example that determines whether the row height changed by more than a small tolerance after AutoFitRows in Aspose.Cells. | Provide a snippet that accesses the IsHeightMatched flag for a row after AutoFitRows to verify the adjustment.
// Common Searches: C# Aspose.Cells get row height before AutoFitRows and after | how to detect row height change after worksheet.AutoFitRows in .NET | Aspose.Cells IsHeightMatched property example for row height verification | compare original and new row heights using Aspose.Cells AutoFitRows
// Tags: Aspose.Cells AutoFitRows row height comparison | C# retrieve row height Aspose.Cells | Aspose.Cells IsHeightMatched usage | verify row height adjustment Aspose.Cells | record original row height points

using System;
using Aspose.Cells;

namespace RowHeightComparisonDemo
{
    // Demonstrates capturing a row's original height, applying AutoFitRows, retrieving the new height, comparing the values with a tolerance, checking the IsHeightMatched flag, and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a cell with long wrapped text to cause row height change
            cells["A1"].PutValue("This is a very long text that should wrap into multiple lines when AutoFitRows is applied.");
            Style style = cells["A1"].GetStyle();
            style.IsTextWrapped = true;               // Enable text wrapping
            cells["A1"].SetStyle(style);

            // Record the original row height (before AutoFitRows)
            double originalHeight = cells.GetRowHeight(0); // height in points
            Console.WriteLine($"Original row height (points): {originalHeight}");

            // Perform AutoFitRows on the entire worksheet
            worksheet.AutoFitRows();

            // Record the new row height (after AutoFitRows)
            double newHeight = cells.GetRowHeight(0);
            Console.WriteLine($"New row height after AutoFitRows (points): {newHeight}");

            // Compare the heights to confirm that an adjustment occurred
            if (Math.Abs(originalHeight - newHeight) > 0.01)
            {
                Console.WriteLine("Row height was adjusted by AutoFitRows.");
            }
            else
            {
                Console.WriteLine("Row height remained unchanged.");
            }

            // Additional check using the IsHeightMatched property
            bool isMatched = worksheet.Cells.Rows[0].IsHeightMatched;
            Console.WriteLine($"IsHeightMatched after AutoFitRows: {isMatched}");

            // Save the workbook (lifecycle rule)
            workbook.Save("RowHeightComparisonDemo.xlsx");
        }
    }
}
