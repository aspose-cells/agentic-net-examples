using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add sample worksheets with various names
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Summary_Q1";

            Worksheet ws2 = workbook.Worksheets.Add("Data");
            Worksheet ws3 = workbook.Worksheets.Add("Annual Summary");
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
            workbook.Save("PaperSizeAssignment.xlsx");
        }
    }
}