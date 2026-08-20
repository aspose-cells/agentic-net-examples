// Title: C# – Set Letter paper size for worksheets with “Summary” in the name, A4 otherwise (Aspose.Cells)
// Description: This .NET example creates a workbook, adds worksheets, then uses a case‑insensitive check on each sheet’s name. Sheets whose name contains “Summary” are assigned the Letter (8.5" × 11") paper size; all other sheets receive A4 (210 mm × 297 mm). The workbook is saved as ConditionalPaperSize.xlsx.
// Keywords: Aspose.Cells | C# | conditional paper size | worksheet PageSetup | Letter paper size | A4 paper size | summary sheet | GitHub example | Aspose.Cells API
// Common Searches: Aspose.Cells set Letter paper size for specific worksheets | C# conditional PageSetup based on worksheet name | How to assign A4 to non‑summary sheets in Aspose.Cells | Aspose.Cells example for paper size by sheet name
// Developer Intent: Apply conditional page‑setup logic so that any worksheet whose name includes "Summary" uses Letter size, while all other worksheets default to A4.
// Use Cases: Generate a US‑oriented summary report that prints on Letter paper while keeping data sheets on A4 for global distribution. | Create a template that automatically applies the correct paper size whenever new worksheets are added, based on naming conventions. | Prepare a multi‑sheet export where printing settings are pre‑configured to avoid manual adjustments before distribution.
// AI Prompts: Write C# code with Aspose.Cells that sets PaperSize to Letter for worksheets whose name contains "Summary" (case‑insensitive) and to A4 for all others. | Show how to modify the loop to also set Landscape orientation for summary sheets while keeping Portrait for the remaining sheets. | Refactor the conditional logic using LINQ to apply the appropriate paper size to every worksheet in a workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    // This .NET example creates a workbook, adds worksheets, then uses a case‑insensitive check on each sheet’s name. Sheets whose name contains “Summary” are assigned the Letter (8.5" × 11") paper size; all other sheets receive A4 (210 mm × 297 mm). The workbook is saved as ConditionalPaperSize.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample worksheets with different names
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            sheet1.Name = "SummaryReport";

            Worksheet sheet2 = workbook.Worksheets.Add("DataSheet");
            Worksheet sheet3 = workbook.Worksheets.Add("AnnualSummary");
            Worksheet sheet4 = workbook.Worksheets.Add("Details");

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // If the worksheet name contains "Summary" (case‑insensitive), set paper size to Letter
                if (ws.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ws.PageSetup.PaperSize = PaperSizeType.PaperLetter; // Letter (8.5" x 11")
                }
                else
                {
                    // Otherwise set paper size to A4
                    ws.PageSetup.PaperSize = PaperSizeType.PaperA4; // A4 (210mm x 297mm)
                }
            }

            // Save the workbook to a file
            workbook.Save("ConditionalPaperSize.xlsx");
        }
    }
}
