// Title: Set Letter paper size for worksheets containing "Summary" and A4 for others using Aspose.Cells C#
// Description: Creates a workbook, adds sample sheets, then iterates through each worksheet. If the sheet name includes the word "Summary" (case‑insensitive) the page setup is set to Letter (8.5" × 11"); otherwise it uses A4 (210 mm × 297 mm). The file is saved as ConditionalPaperSize.xlsx.
// Keywords: Aspose.Cells | C# | conditional paper size | worksheet PageSetup | Letter vs A4 | sheet name filter | US printing format | EU printing format | PageSetup.PaperSize example | programmatic sheet sizing
// Common Searches: Aspose.Cells set paper size based on worksheet name | C# assign Letter size to summary sheets | How to use conditional PageSetup in Aspose.Cells | Change worksheet print format programmatically .NET | Letter and A4 paper size selection Aspose.Cells
// Developer Intent: Apply a case‑insensitive rule that sets Letter paper size for worksheets whose name contains "Summary" and A4 for all remaining sheets.
// Use Cases: Financial reports where summary pages must print on US Letter paper while detailed data uses A4 for international distribution. | Multi‑region workbook generation that automatically matches regional paper standards based on sheet naming conventions. | Automated document preparation for printing services that require different page formats per worksheet.
// AI Prompts: Generate C# code with Aspose.Cells that sets PaperSize to Letter for any worksheet whose Name includes "Summary" (ignore case) and to A4 for all other worksheets. | Show how to loop through a Workbook's Worksheets collection and apply conditional PageSetup.PaperSize based on a substring match. | Explain how to extend the example to use multiple keywords (e.g., "Invoice", "Cover") to assign different paper sizes per worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    // Creates a workbook, adds sample sheets, then iterates through each worksheet. If the sheet name includes the word "Summary" (case‑insensitive) the page setup is set to Letter (8.5" × 11"); otherwise it uses A4 (210 mm × 297 mm). The file is saved as ConditionalPaperSize.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Add sample worksheets for demonstration
            Worksheet ws1 = workbook.Worksheets[0]; // default first sheet
            ws1.Name = "Summary_Q1";

            Worksheet ws2 = workbook.Worksheets.Add("DataSheet");
            Worksheet ws3 = workbook.Worksheets.Add("AnnualSummary2022");
            Worksheet ws4 = workbook.Worksheets.Add("Report");

            // Iterate through all worksheets and set paper size based on name
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // If the worksheet name contains "Summary" (case‑insensitive) use Letter size
                if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter; // 8.5" x 11"
                }
                else
                {
                    // Otherwise use A4 size
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4; // 210mm x 297mm
                }
            }

            // Save the workbook to a file
            workbook.Save("ConditionalPaperSize.xlsx", SaveFormat.Xlsx);
        }
    }
}
