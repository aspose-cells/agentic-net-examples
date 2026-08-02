// Title: C# – Hide Rows 20‑25 in Excel and Convert to PDF using Aspose.Cells
// Description: This example shows how to load an Excel workbook with Aspose.Cells, hide rows 20 through 25 on the first worksheet using the zero‑based HideRows method, and save the modified file directly as a PDF.
// Keywords: Aspose.Cells C# hide rows | HideRows method | Excel to PDF conversion | C# export workbook as PDF | row visibility Aspose.Cells | HideRows(19,6) | Aspose.Cells SaveFormat.Pdf
// Common Searches: Aspose.Cells hide specific rows C# | Convert Excel to PDF after hiding rows | C# HideRows(19,6) example | Omit rows in PDF using Aspose.Cells | Export worksheet to PDF with hidden rows
// Developer Intent: Hide a defined row range in an Excel file and generate a PDF output.
// Use Cases: Prepare client‑ready PDFs that exclude a confidential row block | Create printable reports where particular rows are hidden to improve layout | Batch‑process workbooks to conceal predefined rows before converting each to PDF
// AI Prompts: Generate C# code that uses Aspose.Cells to conceal rows 20‑25 in a worksheet and then save the workbook as a PDF. | Explain the steps to hide a row interval with the HideRows method and export the modified Excel file to PDF in .NET.

using System;
using Aspose.Cells;

// This example shows how to load an Excel workbook with Aspose.Cells, hide rows 20 through 25 on the first worksheet using the zero‑based HideRows method, and save the modified file directly as a PDF.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 20 through 25 (zero‑based start index 19, total 6 rows)
        worksheet.Cells.HideRows(19, 6);

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
