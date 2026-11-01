// Title: Hide column C in an Excel workbook and save as PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file, conceal the third column, and export the worksheet to a PDF using Aspose.Cells in C#. | Programmatically hide column C in a workbook and generate a PDF output with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# hide third column before converting to PDF | How to programmatically hide a column in Excel and export to PDF using Aspose.Cells | C# Aspose.Cells hide column and save workbook as PDF file | Hide specific column in Excel file and generate PDF with Aspose.Cells .NET API
// Tags: Aspose.Cells hide column C | export workbook to PDF Aspose.Cells | Cells.HideColumn method C# | SaveFormat.Pdf conversion Aspose.Cells | modify column visibility before PDF export

using System;
using Aspose.Cells;

// The example loads 'input.xlsx' with Aspose.Cells, accesses the first worksheet, hides column C (zero‑based index 2) using Cells.HideColumn, and then saves the workbook as 'output.pdf' by specifying SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Hide column C (zero‑based index: 2)
        sheet.Cells.HideColumn(2);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
