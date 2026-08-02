// Title: C# – Unhide Columns N‑P (70 pt width) and Export Worksheet to PDF with Aspose.Cells
// Description: Loads an Excel file, unhides columns N through P (indexes 13‑15) on the first worksheet, sets each column’s width to 70 points, and saves the result directly as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# unhide columns | set column width points | export Excel to PDF | column N to P | 70 point column width | worksheet column visibility | PDF conversion Aspose.Cells
// Common Searches: Aspose.Cells unhide columns N P C# | set column width 70 points Aspose.Cells | export Excel worksheet to PDF after unhiding columns | how to reveal hidden columns before PDF conversion Aspose | C# code to unhide specific columns and save as PDF
// Developer Intent: Unhide columns N‑P, apply a 70‑point width, and generate a PDF from the worksheet.
// Use Cases: Generating printable PDF reports that require hidden columns to be visible | Standardizing column width for compliance‑driven PDF layouts | Automating Excel‑to‑PDF conversion in batch jobs where certain columns start hidden | Preparing data sheets for client delivery with specific column dimensions
// AI Prompts: Provide C# Aspose.Cells code that unhides columns 13‑15, sets each column width to 70 points, and saves the workbook as a PDF. | Show how to change column visibility and width in points before converting an Excel worksheet to PDF using Aspose.Cells for .NET. | Explain the steps to load an Excel file, unhide a column range, adjust column width, and export to PDF with Aspose.Cells in C#.

using Aspose.Cells;

// Loads an Excel file, unhides columns N through P (indexes 13‑15) on the first worksheet, sets each column’s width to 70 points, and saves the result directly as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide columns N (index 13) through P (index 15) with a width of 70 points
        worksheet.Cells.UnhideColumns(13, 3, 70);

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
