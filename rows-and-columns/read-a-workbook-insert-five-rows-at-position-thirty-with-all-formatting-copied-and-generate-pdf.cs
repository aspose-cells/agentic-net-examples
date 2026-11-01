// Title: Insert five rows at row 30 with formatting preserved and export the worksheet to PDF using Aspose.Cells for .NET
// AI Prompts: Load an existing Excel workbook, insert five rows beginning at row 30 while copying the source row styles, and save the result as a PDF file using Aspose.Cells in C#. | Using Aspose.Cells, add five new rows at position 30 in the first worksheet, ensure all cell formatting from the original rows is retained, then generate a PDF output. | Programmatically insert multiple rows with style duplication at a specific index and convert the modified workbook to PDF with Aspose.Cells for .NET.
// Common Searches: how to insert rows with style copy in Aspose.Cells C# | Aspose.Cells insert multiple rows at specific index and keep formatting | convert modified Excel workbook to PDF after inserting rows using Aspose.Cells .NET | C# Aspose.Cells insert rows at row 30 and export to PDF | preserve cell formatting when adding rows with Aspose.Cells
// Tags: Aspose.Cells insert rows preserve formatting | C# Aspose.Cells generate PDF from workbook | multiple row insertion Aspose.Cells | copy cell style on row insert Aspose.Cells | save workbook as PDF Aspose.Cells

using Aspose.Cells;
using System;

// The sample loads 'input.xlsx' into a Workbook, inserts five rows at the 30th position (zero‑based index 29) on the first worksheet while copying existing styles, and then saves the modified workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Insert five rows at position 30 (zero‑based index 29) and copy formatting
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.InsertRows(29, 5, true); // true = copy style and formatting

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
