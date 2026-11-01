// Title: Insert a row at index 15 in an Excel worksheet and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Add a single row before Excel row 16 in the first worksheet of input.xlsx and generate a PDF file named output.pdf using Aspose.Cells in C#. | Modify the code to insert three consecutive rows starting at index 15, then convert the updated workbook to PDF.
// Common Searches: Aspose.Cells C# insert row at index 15 before converting to PDF | How to add a row to an Excel file and save as PDF using Aspose.Cells .NET | Insert rows in a worksheet with Aspose.Cells and export the result to PDF | C# code to load Excel, insert row at position 15, and save as PDF with Aspose.Cells
// Tags: add worksheet row Aspose.Cells C# | excel to pdf conversion Aspose.Cells | worksheet row insertion before pdf export | Aspose.Cells zero based row index | save workbook as pdf Aspose.Cells

using Aspose.Cells;
using System;

// The program loads input.xlsx with Aspose.Cells, inserts a row at zero‑based index 15 in the first worksheet, and saves the modified workbook as output.pdf.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new row at index 15 (zero‑based, inserts before the 16th row in Excel)
        sheet.Cells.InsertRows(15, 1);

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
