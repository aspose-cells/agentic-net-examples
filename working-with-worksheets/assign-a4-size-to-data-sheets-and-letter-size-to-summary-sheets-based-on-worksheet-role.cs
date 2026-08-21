// Title: Set A4 for data worksheets and Letter for summary worksheets using Aspose.Cells (.NET)
// Description: C# example that creates a workbook, detects sheet roles from their names, applies PaperSizeType.PaperA4 to data sheets and PaperSizeType.PaperLetter to summary sheets, then saves the file.
// Keywords: Aspose.Cells | C# | paper size | A4 | Letter | worksheet role | page setup | Excel workbook | conditional formatting | print settings
// Common Searches: Aspose.Cells set different paper sizes per worksheet | C# assign A4 to data sheets and Letter to summary sheets | conditional page setup based on sheet name Aspose.Cells | how to change paper size for specific worksheets in .NET
// Developer Intent: Apply role‑based paper sizes—A4 for data sheets, Letter for summary sheets—automatically when generating an Excel workbook.
// Use Cases: Generate printable reports where detailed data tables use A4 and executive summaries use US Letter. | Automate multi‑sheet workbook creation for international partners with locale‑specific page dimensions. | Reduce manual page‑setup steps in batch‑produced Excel files by assigning sizes programmatically.
// AI Prompts: Create a C# Aspose.Cells snippet that sets PaperSizeType.PaperA4 for worksheets whose names contain "Data" and PaperSizeType.PaperLetter for those containing "Summary". | Show how to iterate through all worksheets in a workbook and apply conditional page‑setup settings based on custom role detection. | Explain how to extend the example to also change orientation (portrait/landscape) depending on whether the sheet is a data or summary worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    // C# example that creates a workbook, detects sheet roles from their names, applies PaperSizeType.PaperA4 to data sheets and PaperSizeType.PaperLetter to summary sheets, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample worksheets and assign role via name
            Worksheet dataSheet1 = workbook.Worksheets[0];
            dataSheet1.Name = "Data_January";

            Worksheet dataSheet2 = workbook.Worksheets.Add("Data_February");

            Worksheet summarySheet = workbook.Worksheets.Add("Summary_Report");

            // Iterate through all worksheets and set paper size based on role
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Simple role detection: if name contains "Data" treat as data sheet,
                // otherwise if name contains "Summary" treat as summary sheet.
                if (sheet.Name.IndexOf("Data", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Set A4 size for data sheets
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                }
                else if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Set Letter size for summary sheets
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
                }
                else
                {
                    // Default paper size (optional)
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
                }
            }

            // Save the workbook
            workbook.Save("Workbook_With_PaperSizes.xlsx", SaveFormat.Xlsx);
        }
    }
}
