// Title: C# – Delete Column F with Aspose.Cells and Save Workbook as PDF
// Description: Loads an Excel file, removes column F (zero‑based index 5) from the first worksheet using Aspose.Cells for .NET, updates cross‑sheet references, and saves the result directly as a PDF document.
// Keywords: Aspose.Cells | C# | DeleteColumn | remove column F | Excel to PDF conversion | Workbook.Save PDF | Aspose.Cells .NET | delete column Excel C# | update references Aspose.Cells
// Common Searches: Aspose.Cells delete column C# | How to remove column F from Excel with Aspose.Cells | Convert edited Excel workbook to PDF using Aspose.Cells | DeleteColumn method reference update Aspose.Cells | C# code to delete a column and export PDF
// Developer Intent: Remove column F from an Excel worksheet and export the modified workbook as a PDF using Aspose.Cells for .NET.
// Use Cases: Clean up a template by stripping an unwanted column before generating a PDF report. | Prepare data for distribution when column F contains intermediate calculations that should not appear in the final document. | Batch‑process multiple workbooks to delete a specific column and convert each file to PDF automatically.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete column F (index 5) from the first worksheet and then saves the workbook as a PDF. | Explain the effect of the DeleteColumn method’s second Boolean parameter on formulas and references in other worksheets. | Show how to add error handling for cases where the source file is missing or does not contain column F before performing the deletion and PDF conversion.

using System;
using Aspose.Cells;

// Loads an Excel file, removes column F (zero‑based index 5) from the first worksheet using Aspose.Cells for .NET, updates cross‑sheet references, and saves the result directly as a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Delete column F (zero‑based index 5) from the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteColumn(5, true); // true updates references in other sheets

        // Save the modified workbook as a PDF file
        string outputFile = "output.pdf";
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}
