// Title: Hide column C in an Excel worksheet and save the workbook as a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file, hide column C in the first worksheet, and generate a PDF output with Aspose.Cells in C#. | Using Aspose.Cells for .NET, conceal a specific column in an Excel workbook and export the modified sheet to PDF.
// Common Searches: Aspose.Cells C# hide column C before exporting to PDF | how to hide a specific column in Excel and save as PDF using Aspose.Cells | C# code to hide column in worksheet and convert workbook to PDF with Aspose.Cells | export modified Excel sheet to PDF after hiding columns Aspose.Cells .NET | Aspose.Cells hide column then save as PDF example
// Tags: hide column Aspose.Cells C# | export worksheet to PDF Aspose.Cells | Aspose.Cells column visibility manipulation | PDF conversion after column hide Aspose.Cells | Aspose.Cells hide specific column before PDF save

using System;
using Aspose.Cells;

// The example loads input.xlsx, hides column C (index 2) in the first worksheet, and saves the workbook as output.pdf in PDF format using Aspose.Cells for .NET.
class HideColumnAndSavePdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputFile = "output.pdf";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide column C (zero‑based index 2)
        worksheet.Cells.HideColumn(2);

        // Save the modified workbook as PDF
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}
