// Title: Hide rows 21‑25 in an Excel worksheet and export as PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file, conceals rows 21‑25 on the first worksheet using Aspose.Cells, and saves the workbook as a PDF. | Create a C# program that uses Aspose.Cells to hide a specific row range before converting the workbook to PDF format.
// Common Searches: Aspose.Cells C# hide rows 21 to 25 before PDF export | How to hide a range of rows in Excel with Aspose.Cells and generate PDF | C# example for hiding rows in worksheet and saving as PDF using Aspose.Cells | Convert Excel to PDF while keeping rows hidden using Aspose.Cells .NET
// Tags: Worksheet.Cells.HideRows method C# | Aspose.Cells PDF export hidden rows | C# convert Excel to PDF Aspose.Cells | row visibility control Aspose.Cells | Excel to PDF conversion Aspose.Cells C#

using System;
using Aspose.Cells;

namespace HideRowsAndSavePdf
{
    // Loads input.xlsx, uses Worksheet.Cells.HideRows to conceal rows 21‑25 on the first sheet, then saves the workbook as output.pdf, preserving the hidden rows in the PDF output.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide rows 21 to 25 (zero‑based index: 20 to 24)
            // HideRows(startRowIndex, totalRows)
            worksheet.Cells.HideRows(20, 5);

            // Save the workbook as PDF; hidden rows will be reflected in the output
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
