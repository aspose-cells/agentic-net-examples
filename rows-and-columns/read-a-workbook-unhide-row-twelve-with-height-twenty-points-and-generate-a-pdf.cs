// Title: Aspose.Cells C# – Unhide Row 12, Set Height to 20 pt, Export Workbook as PDF
// Description: Load an Excel file with Aspose.Cells for .NET, unhide row 12 (index 11), set its height to 20 points, and save the workbook directly to a PDF document.
// Keywords: Aspose.Cells unhide row C# | set row height points | export Excel to PDF .NET | unhide specific row Aspose.Cells | row height PDF conversion
// Common Searches: Aspose.Cells how to unhide a row before PDF export | C# set row height in points with Aspose.Cells | convert Excel to PDF after unhiding rows | unhide row 12 Aspose.Cells .NET
// Developer Intent: Unhide row 12, assign a 20‑point height, and generate a PDF from the workbook using Aspose.Cells for .NET.
// Use Cases: Prepare printable reports where hidden rows must appear with a fixed height before PDF conversion. | Standardize layout of Excel templates by revealing specific rows and defining their height prior to distribution as PDF. | Automate batch processing of workbooks to ensure required rows are visible and uniformly sized before each file is saved as PDF.
// AI Prompts: Generate C# code with Aspose.Cells that unhides row 15, sets its height to 25 pt, and saves the workbook as a PDF. | Create a reusable method accepting a worksheet, row index, height, and PDF path to unhide the row, adjust its height, and export to PDF. | Explain how to unhide multiple rows and assign custom heights before converting an Excel workbook to PDF using Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel file with Aspose.Cells for .NET, unhide row 12 (index 11), set its height to 20 points, and save the workbook directly to a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide row 12 (zero‑based index 11) and set its height to 20 points
        worksheet.Cells.UnhideRow(11, 20);

        // Save the workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
