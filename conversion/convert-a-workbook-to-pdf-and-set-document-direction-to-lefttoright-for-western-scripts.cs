// Title: C# – Convert Aspose.Cells Workbook to PDF with Left‑to‑Right Layout for Western Scripts
// Description: Creates a new Workbook, disables the right‑to‑left view on each worksheet (DisplayRightToLeft = false), adds optional sample data, and saves the file as a PDF using Workbook.Save with SaveFormat.Pdf, ensuring the output follows Western reading order.
// Keywords: Aspose.Cells PDF conversion C# | DisplayRightToLeft false | left to right Excel PDF | Workbook.Save PDF Aspose | Western script PDF export | C# Excel to PDF Aspose.Cells
// Common Searches: Aspose.Cells export Excel to PDF left to right | disable right to left layout before PDF conversion Aspose | C# set worksheet direction left to right Aspose.Cells | how to force left‑to‑right PDF output with Aspose.Cells
// Developer Intent: Produce a PDF from an Excel workbook while enforcing a left‑to‑right page direction.
// Use Cases: Generate PDF reports for English‑speaking users where RTL layout would be incorrect. | Automate batch conversion of multiple workbooks, ensuring each PDF follows Western reading order. | Create printable invoices or dashboards that contain dates and text and must appear left‑to‑right.
// AI Prompts: Show C# code that loads an Excel file with Aspose.Cells, sets DisplayRightToLeft to false for all sheets, and saves it as PDF. | Explain how to configure Aspose.Cells PDF conversion to enforce left‑to‑right document direction. | Give a step‑by‑step example of converting a dynamically generated workbook to PDF with Western layout using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfConversion
{
    // Creates a new Workbook, disables the right‑to‑left view on each worksheet (DisplayRightToLeft = false), adds optional sample data, and saves the file as a PDF using Workbook.Save with SaveFormat.Pdf, ensuring the output follows Western reading order.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure left‑to‑right display for all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.DisplayRightToLeft = false;
            }

            // Add sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Left‑to‑Right Example");
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // Save the workbook as PDF using the provided Save(string, SaveFormat) rule
            workbook.Save("output.pdf", SaveFormat.Pdf);

            Console.WriteLine("Workbook successfully converted to PDF with left‑to‑right direction.");
        }
    }
}
