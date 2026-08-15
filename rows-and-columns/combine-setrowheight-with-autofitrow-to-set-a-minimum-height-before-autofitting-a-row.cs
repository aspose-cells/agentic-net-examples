// Title: Set Minimum Row Height and AutoFit Row with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes wrapped text to a cell, applies a baseline height using SetRowHeight, auto‑fits the row with AutoFitRow, and guarantees the final height does not fall below the defined threshold before saving.
// Keywords: Aspose.Cells SetRowHeight | Aspose.Cells AutoFitRow | C# minimum row height | wrap text row height Aspose | Excel row height control .NET | Aspose.Cells row height example
// Common Searches: Aspose.Cells set row height then autofit | C# enforce minimum row height in Excel with Aspose | prevent AutoFitRow from reducing row height | set minimum row height before AutoFitRow Aspose.Cells | wrap text and auto‑fit row height C# Aspose
// Developer Intent: Define a baseline height, auto‑fit the row based on its content, and ensure the height stays at or above that baseline.
// Use Cases: Generating reports where wrapped text rows must be at least 30 pt high for readability. | Building templates that automatically enforce a minimum row height after populating data. | Exporting data to Excel while guaranteeing specific rows never shrink below a set height.
// AI Prompts: Show C# code using Aspose.Cells that sets a minimum row height, enables text wrapping, auto‑fits the row, and re‑applies the minimum if needed. | Provide an Aspose.Cells example that combines SetRowHeight and AutoFitRow to keep row height above a threshold after fitting content. | Write a .NET snippet that writes long wrapped text to a cell, applies a baseline row height, calls worksheet.AutoFitRow, and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, writes wrapped text to a cell, applies a baseline height using SetRowHeight, auto‑fits the row with AutoFitRow, and guarantees the final height does not fall below the defined threshold before saving.
    public class SetRowHeightWithAutoFitDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Row index to work with
            int rowIndex = 0;

            // Define a minimum row height (in points)
            double minHeight = 30.0; // Example: 30 points

            // Put some sample text that may require a larger height
            cells["A1"].PutValue("This is a very long text that should cause the row to expand when auto‑fitted.");
            Style style = cells["A1"].GetStyle();
            style.IsTextWrapped = true; // Enable wrapping to affect row height
            cells["A1"].SetStyle(style);

            // Set the minimum row height before auto‑fitting (rule: SetRowHeight)
            cells.SetRowHeight(rowIndex, minHeight);

            // Auto‑fit the row based on its content (rule: AutoFitRow)
            worksheet.AutoFitRow(rowIndex);

            // Ensure the row height is not less than the defined minimum
            double finalHeight = cells.GetRowHeight(rowIndex);
            if (finalHeight < minHeight)
            {
                cells.SetRowHeight(rowIndex, minHeight);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("SetRowHeightWithAutoFitDemo.xlsx");
        }
    }
}
