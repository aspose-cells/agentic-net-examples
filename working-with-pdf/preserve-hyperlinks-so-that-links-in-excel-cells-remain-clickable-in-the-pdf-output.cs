// Title: C# – Preserve Hyperlinks When Converting Excel to PDF with Aspose.Cells
// Description: Demonstrates how to add a hyperlink to an Excel cell and export the workbook to PDF using Aspose.Cells for .NET, keeping the link clickable in the resulting PDF file.
// Keywords: Aspose.Cells hyperlink PDF | C# export Excel to PDF with links | clickable links in PDF from Excel | .NET preserve Excel hyperlinks PDF | global PDF export Aspose.Cells
// Common Searches: keep Excel hyperlinks after PDF conversion C# | Aspose.Cells export PDF retain links | how to make PDF hyperlinks from Excel using .NET | C# code to preserve cell links in PDF
// Developer Intent: Export an Excel workbook to PDF while maintaining active hyperlinks in the PDF.
// Use Cases: Create PDF reports from spreadsheets where embedded URLs must stay interactive. | Generate marketing brochures from Excel templates that contain product or support links. | Distribute financial models or dashboards as PDFs with reference links that remain functional.
// AI Prompts: Write C# code with Aspose.Cells that converts an Excel file to PDF and keeps all cell hyperlinks clickable. | Explain the hyperlink preservation mechanism in Aspose.Cells PDF export and any settings that affect it. | Show how to add multiple hyperlinks to different cells and export the workbook to a PDF with active links using Aspose.Cells for .NET.

using Aspose.Cells;
using System;

// Demonstrates how to add a hyperlink to an Excel cell and export the workbook to PDF using Aspose.Cells for .NET, keeping the link clickable in the resulting PDF file.
class PreserveHyperlinksPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put display text into a cell
        worksheet.Cells["A1"].PutValue("Visit Aspose");

        // Add a hyperlink to the cell (A1)
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Save the workbook as PDF; hyperlinks remain clickable in the PDF output
        workbook.Save("HyperlinksPreserved.pdf");
    }
}
