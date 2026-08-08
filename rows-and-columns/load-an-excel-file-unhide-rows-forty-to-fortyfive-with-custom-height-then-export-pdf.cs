// Title: C# – Unhide Rows 40‑45, Set Custom Height, and Export Excel to PDF with Aspose.Cells
// Description: Loads an existing workbook, unhides rows 40 through 45, applies a 20‑point row height, and saves the worksheet as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel to PDF | UnhideRows method | custom row height | set row height Aspose | export workbook as PDF | Excel row visibility | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells unhide rows 40 to 45 C# | set custom row height and export to PDF using Aspose | how to make hidden rows visible before PDF conversion in .NET | C# code to unhide specific rows and save Excel as PDF | Aspose.Cells UnhideRows parameters example
// Developer Intent: Unhide rows 40‑45, assign a specific height, and generate a PDF from the modified workbook.
// Use Cases: Prepare a financial statement by revealing hidden rows and adjusting layout before creating a printable PDF. | Standardize row heights in a report template prior to distribution as a PDF document. | Automate the cleanup of hidden rows in batch‑processed spreadsheets before archiving them in PDF format.
// AI Prompts: Write C# code that uses Aspose.Cells to unhide rows 40‑45, set each row to 20 points high, and save the file as a PDF. | Explain the three parameters of the Cells.UnhideRows method and how they influence PDF output. | Add comprehensive error handling to the sample that loads an Excel file, modifies row visibility, and exports to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an existing workbook, unhides rows 40 through 45, applies a 20‑point row height, and saves the worksheet as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Unhide rows 40‑45 (zero‑based index) and set a custom height (e.g., 20 points)
        cells.UnhideRows(40, 6, 20.0);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
