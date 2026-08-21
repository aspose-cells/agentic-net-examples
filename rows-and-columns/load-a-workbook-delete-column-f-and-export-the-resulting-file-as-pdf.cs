// Title: Delete Column F in an Excel workbook and export it as PDF using Aspose.Cells for .NET
// Description: Load an existing workbook, remove column F (zero‑based index 5) from the first worksheet with Aspose.Cells, and save the modified file directly as a PDF.
// Keywords: Aspose.Cells delete column | remove Excel column .NET | export Excel to PDF Aspose | column deletion before PDF conversion | C# Aspose.Cells PDF output
// Common Searches: Aspose.Cells delete column F and save as PDF | C# remove specific column from Excel then convert to PDF | how to export modified worksheet to PDF with Aspose.Cells
// Developer Intent: Strip column F from an Excel sheet and generate a PDF of the result.
// Use Cases: Omit confidential data before sharing a financial spreadsheet as PDF. | Create clean printable invoices by removing internal notes columns. | Archive a dataset after discarding unnecessary columns and saving it as PDF.
// AI Prompts: Generate C# code that deletes column G from the second worksheet and saves the workbook as a PDF using Aspose.Cells. | Explain how to remove multiple columns from a worksheet and then export the workbook to PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an existing workbook, remove column F (zero‑based index 5) from the first worksheet with Aspose.Cells, and save the modified file directly as a PDF.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete column F (zero‑based column index 5)
        sheet.Cells.DeleteColumn(5);

        // Save the modified workbook as a PDF file
        string outputPath = "output.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
