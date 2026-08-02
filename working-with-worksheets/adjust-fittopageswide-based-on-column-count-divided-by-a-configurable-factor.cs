// Title: Aspose.Cells .NET – Dynamically Set FitToPagesWide Using Column Count and a Configurable Factor
// Description: This example creates a workbook, fills it with sample data, determines the last used column, calculates the required number of pages wide by dividing the column count by a user‑defined factor (rounded up, minimum 1), applies the value to worksheet.PageSetup.FitToPagesWide, sets FitToPagesTall to auto, and saves the file.
// Keywords: Aspose.Cells FitToPagesWide dynamic | C# calculate pages wide from columns | configure print scaling Aspose.Cells | worksheet page setup factor .NET | auto fit columns to pages Aspose
// Common Searches: Aspose.Cells set FitToPagesWide by column count | C# calculate pages wide for printing Excel | dynamic page width factor Aspose.Cells | auto adjust worksheet print scaling .NET
// Developer Intent: Compute and assign worksheet.PageSetup.FitToPagesWide based on total used columns divided by a configurable factor, ensuring at least one page.
// Use Cases: Print reports where a fixed number of columns should appear on each page. | Generate Excel files that automatically adapt page width when column counts vary. | Create printable workbooks with consistent column grouping across pages.
// AI Prompts: Generate C# code with Aspose.Cells that sets FitToPagesWide from the column count and a user‑specified factor, rounding up and enforcing a minimum of one page. | Explain how to retrieve the last used column in an Aspose.Cells worksheet and calculate the pages‑wide value for printing. | Show an example that sets FitToPagesTall to auto while FitToPagesWide is derived from column count and a configurable divisor.

using System;
using Aspose.Cells;

namespace AsposeCellsFitToPagesWideDemo
{
    // This example creates a workbook, fills it with sample data, determines the last used column, calculates the required number of pages wide by dividing the column count by a user‑defined factor (rounded up, minimum 1), applies the value to worksheet.PageSetup.FitToPagesWide, sets FitToPagesTall to auto, and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Configurable factor to divide the column count
            int factor = 5; // Example: fit every 5 columns onto one page

            // ---------- Create ----------
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (e.g., 20 columns)
            for (int col = 0; col < 20; col++)
            {
                for (int row = 0; row < 10; row++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ---------- Compute FitToPagesWide ----------
            // MaxColumn returns the zero‑based index of the last used column
            int lastUsedColumnIndex = worksheet.Cells.MaxColumn;
            int totalColumns = lastUsedColumnIndex + 1; // Convert to count

            // Calculate pages wide, rounding up to ensure all columns fit
            int pagesWide = (int)Math.Ceiling((double)totalColumns / factor);
            if (pagesWide < 1) pagesWide = 1; // Ensure at least one page

            // Set page setup: fit to calculated pages wide, height adjusts automatically
            worksheet.PageSetup.FitToPagesWide = pagesWide;
            worksheet.PageSetup.FitToPagesTall = 0; // Let height scale automatically

            // ---------- Save ----------
            // Save the workbook to a file
            workbook.Save("FitToPagesWideAdjusted.xlsx");
        }
    }
}
