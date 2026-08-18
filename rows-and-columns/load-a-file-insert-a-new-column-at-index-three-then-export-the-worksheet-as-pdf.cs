// Title: C# – Insert a Column at Index 3 and Save Worksheet as PDF using Aspose.Cells
// Description: Loads input.xlsx, accesses the first sheet, inserts a new column at the fourth position (zero‑based index 3), and writes the workbook to output.pdf with PdfSaveOptions.
// Keywords: Aspose.Cells | C# | InsertColumn | Excel to PDF | PdfSaveOptions | modify worksheet columns | export Excel as PDF | zero based column index
// Common Searches: Aspose.Cells insert column at position 4 C# | Save Excel worksheet as PDF after adding column Aspose | How to add a blank column in .NET Excel file and export to PDF | Insert column zero based index Aspose.Cells example | Convert modified workbook to PDF using Aspose.Cells
// Developer Intent: Add a column at the fourth position in an Excel sheet and generate a PDF file.
// Use Cases: Prepare a printable PDF report after inserting a new data column. | Adjust column layout in a financial template before distribution as PDF. | Automate column insertion in batch processing and output PDFs. | Create a PDF version of a spreadsheet after structural changes.
// AI Prompts: Write C# code with Aspose.Cells that inserts a column at zero‑based index 3 and saves the first worksheet as a PDF. | Show how to use PdfSaveOptions to export a workbook to PDF after adding a column. | Explain the steps to modify column structure in an Excel file with Aspose.Cells and generate a PDF without affecting other sheets.

using System;
using Aspose.Cells;

// Loads input.xlsx, accesses the first sheet, inserts a new column at the fourth position (zero‑based index 3), and writes the workbook to output.pdf with PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a new column at index 3 (zero‑based, i.e., the fourth column)
        worksheet.Cells.InsertColumn(3);

        // Export the worksheet (entire workbook) to PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);
    }
}
