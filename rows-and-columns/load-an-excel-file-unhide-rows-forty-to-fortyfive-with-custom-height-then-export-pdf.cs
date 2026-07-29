// Title: C# – Unhide Rows 40‑45, Set Custom Height, and Export to PDF with Aspose.Cells
// Description: Loads an Excel workbook, unhides rows 40‑45 on the first worksheet, sets each row’s height to 15 points, and saves the file as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# unhide rows | set row height Aspose.Cells | export Excel to PDF .NET | unhide rows 40-45 | custom row height PDF | Aspose.Cells UnhideRows method | Excel to PDF conversion C# | row visibility Aspose.Cells
// Common Searches: Aspose.Cells unhide rows and export to PDF | C# set row height before PDF conversion | How to reveal hidden rows in Excel with Aspose.Cells | UnhideRows method example C# | Export Excel worksheet to PDF after changing row visibility
// Developer Intent: The developer wants to programmatically make rows 40‑45 visible, assign a specific height, and generate a PDF from the updated workbook.
// Use Cases: Expose hidden rows in a financial statement before creating a printable PDF for stakeholders. | Adjust row heights to improve layout when exporting an invoice template to PDF. | Prepare a grade‑book Excel file by unhiding selected rows and delivering the final version as a PDF report. | Generate a marketing deck where certain rows are hidden during editing but need to appear in the final PDF.
// AI Prompts: Generate C# code that uses Aspose.Cells to unhide rows 40‑45, set each row height to 15 points, and save the workbook as a PDF. | Explain the parameters of the UnhideRows method in Aspose.Cells and how to specify a custom height for the rows. | Provide a step‑by‑step tutorial for loading an Excel file, changing row visibility and height, and exporting the result to PDF with Aspose.Cells for .NET. | Troubleshoot why rows remain hidden after calling UnhideRows in Aspose.Cells and suggest fixes.

using System;
using Aspose.Cells;

// Loads an Excel workbook, unhides rows 40‑45 on the first worksheet, sets each row’s height to 15 points, and saves the file as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Unhide rows 40 to 45 (zero‑based index 39, total 6 rows) and set a custom height (e.g., 15 points)
        workbook.Worksheets[0].Cells.UnhideRows(39, 6, 15.0);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
