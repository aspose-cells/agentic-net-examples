// Title: Aspose.Cells for .NET – Delete Blank Rows and Save Workbook as PDF (C#)
// Description: Loads an Excel workbook, removes every completely empty row from the first worksheet using DeleteBlankRows, and saves the cleaned workbook directly as a PDF with SaveFormat.Pdf.
// Keywords: Aspose.Cells | C# DeleteBlankRows | remove empty rows Excel | export Excel to PDF .NET | compact worksheet PDF | Aspose.Cells SaveFormat.Pdf | Excel blank rows removal
// Common Searches: Aspose.Cells delete blank rows C# | How to remove empty rows before PDF export using Aspose.Cells | C# code to convert Excel to PDF after cleaning rows | DeleteBlankRows method example | Save workbook as PDF with Aspose.Cells .NET
// Developer Intent: Delete every completely empty row from the first worksheet of an Excel file and generate a PDF version of the cleaned workbook using Aspose.Cells for .NET.
// Use Cases: Sanitize data imports by stripping blank rows prior to PDF reporting. | Create space‑efficient PDF archives of spreadsheets. | Integrate into batch jobs that convert multiple Excel files to compact PDFs.
// AI Prompts: Write C# code that iterates through all worksheets, deletes blank rows in each, and saves each as a separate PDF file using Aspose.Cells. | Provide a snippet that logs the count of removed rows before PDF conversion. | Show how to add robust error handling for file I/O and PDF generation with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, removes every completely empty row from the first worksheet using DeleteBlankRows, and saves the cleaned workbook directly as a PDF with SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete all blank rows in the worksheet
        worksheet.Cells.DeleteBlankRows();

        // Save the compacted workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
