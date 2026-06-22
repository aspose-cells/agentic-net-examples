using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsThemeAudit
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be audited
            string workbookPath = "input.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare a list to hold addresses of cells using the Light1 (Background1) theme color
            List<string> light1Cells = new List<string>();

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    // Retrieve the cell; may be null if the cell is empty
                    Cell cell = cells[row, col];
                    if (cell == null) continue;

                    // Get the cell's style
                    Style style = cell.GetStyle();

                    // Check if the background theme color is set and matches Background1 (Light1)
                    ThemeColor bgTheme = style.BackgroundThemeColor;
                    if (bgTheme != null && bgTheme.ColorType == ThemeColorType.Background1)
                    {
                        // Add the cell address (e.g., "A1") to the result list
                        light1Cells.Add(cell.Name);
                    }
                }
            }

            // Output the audit results
            Console.WriteLine("Cells using the Light1 (Background1) theme background:");
            foreach (string address in light1Cells)
            {
                Console.WriteLine(address);
            }

            // Optionally, save the workbook after audit (uses the provided save rule)
            // workbook.Save("audited_output.xlsx");
        }
    }
}