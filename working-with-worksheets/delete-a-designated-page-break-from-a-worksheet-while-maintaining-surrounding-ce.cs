using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDeletion
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet by index/name)
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Define the page break to delete.
            // Example: delete the horizontal page break that starts at row 9 (zero‑based index)
            // If you prefer to specify by cell name, use e.g. "B10" and convert to row index.
            // ------------------------------------------------------------
            int targetRow = 9; // zero‑based row index of the page break to remove

            // Find the index of the horizontal page break that matches the target row
            int breakIndex = -1;
            for (int i = 0; i < worksheet.HorizontalPageBreaks.Count; i++)
            {
                HorizontalPageBreak hBreak = worksheet.HorizontalPageBreaks[i];
                if (hBreak.Row == targetRow)
                {
                    breakIndex = i;
                    break;
                }
            }

            // If the page break exists, remove it while preserving cell formatting/layout
            if (breakIndex >= 0)
            {
                worksheet.HorizontalPageBreaks.RemoveAt(breakIndex);
                Console.WriteLine($"Horizontal page break at row {targetRow} removed.");
            }
            else
            {
                Console.WriteLine($"No horizontal page break found at row {targetRow}.");
            }

            // ------------------------------------------------------------
            // Save the modified workbook (replace with your desired output path)
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}