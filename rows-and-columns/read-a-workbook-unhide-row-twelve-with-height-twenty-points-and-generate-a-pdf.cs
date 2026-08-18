// Title: C# – Unhide Row 12, Set Height to 20 pt, and Export Workbook to PDF with Aspose.Cells
// Description: Load an Excel file, unhide row 12 (index 11), set its height to 20 points, and save the workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# unhide row | set row height Aspose.Cells | Excel to PDF conversion .NET | row visibility Aspose.Cells | export workbook as PDF | unhide specific row PDF
// Common Searches: Aspose.Cells unhide row 12 and export to PDF | C# set row height before PDF conversion Aspose.Cells | how to make hidden row visible in Excel and save as PDF using Aspose | Aspose.Cells change row visibility and height for PDF output
// Developer Intent: Make row 12 visible with a 20‑point height and generate a PDF from the workbook.
// Use Cases: Show a hidden header row, adjust its size, and create a printable PDF report. | Ensure a concealed data row appears with the correct height in a PDF invoice. | Automate Excel‑to‑PDF conversion while programmatically controlling row visibility and dimensions.
// AI Prompts: Generate C# code with Aspose.Cells that unhides row 12, sets its height to 20 points, and saves the workbook as a PDF. | Explain how to change row visibility and height in Aspose.Cells before exporting to PDF, including required namespaces and method calls.

using System;
using Aspose.Cells;

// Load an Excel file, unhide row 12 (index 11), set its height to 20 points, and save the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide row 12 (zero‑based index 11) and set its height to 20 points
        worksheet.Cells.UnhideRow(11, 20);

        // Save the workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
