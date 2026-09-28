// Title: How to enforce a minimum row height then auto‑fit a row with wrapped text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that defines a lower bound for a worksheet row's height, enables text wrapping, calls Worksheet.AutoFitRow, and restores the lower bound if the auto‑fit result is smaller, using Aspose.Cells. | Write a .NET example that sets a 20‑point lower limit for row 0, inserts a long wrapped string into cell A1, auto‑fits the row, and guarantees the final height never drops below that limit.
// Common Searches: Aspose.Cells C# set minimum row height before AutoFitRow | keep row height from shrinking after AutoFitRow in .NET | auto fit row with text wrap and minimum height using Aspose.Cells | C# Aspose.Cells enforce row height lower bound when auto fitting | how to prevent Aspose.Cells AutoFitRow from reducing row height
// Tags: set row height minimum limit Aspose.Cells C# | auto‑fit row respecting minimum limit Aspose.Cells | text wrap influences row height Aspose.Cells | prevent row height shrinkage Aspose.Cells | Aspose.Cells worksheet row height management .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, applies a 20‑point minimum height to row 0, adds long wrapped text to cell A1, auto‑fits the row, checks the resulting height, and restores the minimum height if the auto‑fit makes the row shorter, then saves the file as SetRowHeightWithAutoFitDemo.xlsx.
    public class SetRowHeightWithAutoFitDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Row index to work with
                int rowIndex = 0;

                // Set a minimum row height (in points)
                double minHeight = 20.0; // points
                cells.SetRowHeight(rowIndex, minHeight);

                // Add long text that may require a larger height
                cells["A1"].PutValue("This is a very long piece of text that should cause the row to expand when auto‑fitted.");
                Style style = cells["A1"].GetStyle();
                style.IsTextWrapped = true; // Enable wrapping so the text affects row height
                cells["A1"].SetStyle(style);

                // Auto‑fit the row based on its content
                worksheet.AutoFitRow(rowIndex);

                // Ensure the row height is not less than the minimum height
                double actualHeight = cells.GetRowHeight(rowIndex);
                if (actualHeight < minHeight)
                {
                    cells.SetRowHeight(rowIndex, minHeight);
                }

                // Save the workbook
                workbook.Save("SetRowHeightWithAutoFitDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRowHeightWithAutoFitDemo.Run();
        }
    }
}
