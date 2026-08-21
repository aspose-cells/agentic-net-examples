// Title: Hide Column C in Excel and Export to PDF with Aspose.Cells for .NET (C#)
// Description: Load an Excel workbook using Aspose.Cells for .NET, hide column C (index 2) on the first worksheet, and save the result directly as a PDF file.
// Keywords: Aspose.Cells hide column C | C# hide Excel column | export Excel to PDF Aspose.Cells | Aspose.Cells PDF conversion | hide column before PDF export | Aspose.Cells .NET | Excel column visibility | PDF export C#
// Common Searches: C# hide column C Aspose.Cells | Aspose.Cells hide Excel column and save as PDF | How to hide a column in Excel using Aspose.Cells .NET | Export hidden columns Excel to PDF with Aspose.Cells | Aspose.Cells PDF conversion with hidden columns
// Developer Intent: Hide column C in an Excel worksheet and generate a PDF file.
// Use Cases: Produce printable PDFs that exclude confidential data located in column C. | Create clean‑layout PDF reports by hiding unnecessary columns before conversion. | Automate batch processing to hide specific columns across multiple workbooks prior to PDF export.
// AI Prompts: Generate C# code to hide multiple columns (e.g., B, D) and export to PDF with custom page size using Aspose.Cells. | Explain how to hide a column by its letter versus index and adjust PDF export options such as image quality and orientation. | Show how to hide column C, set PDF save options (e.g., compliance, compression), and save the workbook in one step.

using System;
using Aspose.Cells;

// Load an Excel workbook using Aspose.Cells for .NET, hide column C (index 2) on the first worksheet, and save the result directly as a PDF file.
class HideColumnAndSavePdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputFile = "output.pdf";

        // Load the workbook (lifecycle create/load)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide column C (zero‑based index 2)
        worksheet.Cells.HideColumn(2);

        // Save the modified workbook as PDF (lifecycle save)
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}
