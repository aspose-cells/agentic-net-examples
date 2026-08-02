// Title: Unhide columns H‑J, set 50‑point width, and export to PDF with Aspose.Cells for .NET (C#)
// Description: This C# sample loads an Excel workbook, unhides columns H through J, sets each column’s width to 50 points, and saves the workbook directly as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | unhide columns | column width points | UnhideColumns method | Excel to PDF | PDF export | worksheet manipulation | set column width Aspose
// Common Searches: Aspose.Cells unhide columns H J C# | set column width in points Aspose.Cells | export Excel workbook to PDF after changing column visibility | C# code to unhide multiple columns and save as PDF | how to use UnhideColumns method Aspose.Cells
// Developer Intent: Reveal hidden columns H‑J, apply a 50‑point width, and generate a PDF from the modified workbook.
// Use Cases: Prepare a financial report by making hidden columns visible and uniformly wide before creating a printable PDF. | Adjust layout of a marketing dashboard worksheet, then deliver the final design as a PDF to stakeholders. | Automate data‑entry form cleanup by unhiding specific columns, setting consistent widths, and exporting the result for archival.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, unhides columns H‑J, sets each column width to 50 points, and saves the workbook as a PDF. | Explain the parameters of the UnhideColumns method in Aspose.Cells and how to specify column width in points before PDF conversion. | Show how to modify column visibility and width in a worksheet without affecting other sheets, then export the workbook to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// This C# sample loads an Excel workbook, unhides columns H through J, sets each column’s width to 50 points, and saves the workbook directly as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide columns H (index 7) through J (index 9) – total 3 columns – and set their width to 50 points
        worksheet.Cells.UnhideColumns(7, 3, 50);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
