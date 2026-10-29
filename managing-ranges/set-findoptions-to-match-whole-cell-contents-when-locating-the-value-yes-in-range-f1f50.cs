// Title: Use Aspose.Cells for .NET to find exact "Yes" values in column F (F1:F50) and apply a green highlight
// Common Searches: Aspose.Cells C# find exact text "Yes" in column F rows 1 to 50 | How to use FindOptions to match whole cell content in Aspose.Cells .NET | Apply background color to cells that equal "Yes" using Aspose.Cells C# | Highlight cells with value Yes in a specific range with Aspose.Cells | C# code to search column F for whole‑cell match and style the result in Aspose.Cells
// Tags: FindOptions whole cell match Aspose.Cells | exact text search column F Aspose.Cells | apply background style to matched cells Aspose.Cells | range F1:F50 cell value lookup Aspose.Cells | highlight Excel cells based on exact value Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // The program creates a new workbook, iterates rows 1‑50 in column F, checks each cell for the exact string "Yes", applies a light‑green solid background style to matching cells, and saves the workbook as Result.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target column (F) and row range (1‑50)
                int targetColumn = 5; // zero‑based index for column F
                int startRow = 0;     // row 1 (zero‑based)
                int endRow = 49;      // row 50 (zero‑based)

                // Iterate through the specified range and look for the exact value "Yes"
                for (int row = startRow; row <= endRow; row++)
                {
                    Cell cell = worksheet.Cells[row, targetColumn];
                    if (cell != null && cell.Value != null && cell.Value.ToString() == "Yes")
                    {
                        // Apply a visual highlight to the matching cell
                        Style style = cell.GetStyle();
                        style.ForegroundColor = Color.LightGreen;
                        style.Pattern = BackgroundType.Solid;
                        cell.SetStyle(style);
                        // If only the first match is needed, break here
                        // break;
                    }
                }

                // Save the workbook
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
