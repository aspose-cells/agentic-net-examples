// Title: C# – Merge Cells C6:E7, Apply Bold Font, and Save as XLSX with Aspose.Cells
// Description: This example demonstrates how to load an existing XLSX workbook using Aspose.Cells for .NET, merge the range C6:E7 on the first worksheet, apply a bold font style to the merged area, and save the result as a new XLSX file.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | merge cells | C6:E7 | bold font | cell styling | save workbook as XLSX | load Excel file | Excel automation
// Common Searches: Aspose.Cells merge cells C6 to E7 C# | How to apply bold font to merged cells with Aspose.Cells .NET | C# code to merge a range and save workbook as XLSX using Aspose.Cells | Aspose.Cells example: merge cells and set style | Load, modify, and save Excel file with Aspose.Cells C#
// Developer Intent: The developer needs to merge cells C6:E7, make the content bold, and export the workbook as an XLSX file using Aspose.Cells in C#.
// Use Cases: Create a centered report title that spans columns C‑E with bold formatting. | Design a multi‑column table header that appears bold in generated spreadsheets. | Build a template where a merged cell serves as a highlighted section label.
// AI Prompts: Generate C# code with Aspose.Cells to merge cells A1:D2, set italic style, and save as XLSX. | Explain how to add background color, alignment, and border to a merged cell range using Aspose.Cells for .NET. | Provide a step‑by‑step guide to merge a cell range, apply multiple style attributes (font, color, alignment), and export the workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// This example demonstrates how to load an existing XLSX workbook using Aspose.Cells for .NET, merge the range C6:E7 on the first worksheet, apply a bold font style to the merged area, and save the result as a new XLSX file.
class Program
{
    static void Main()
    {
        // Load the workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells C6:E7.
        // C = column index 2, 6 = row index 5 (zero‑based).
        // Total rows = 2 (rows 6 and 7), total columns = 3 (C, D, E).
        cells.Merge(5, 2, 2, 3);

        // Apply bold font to the merged cell (reference the upper‑left cell C6).
        Style style = cells[5, 2].GetStyle();
        style.Font.IsBold = true;
        cells[5, 2].SetStyle(style);

        // Save the modified workbook as XLSX.
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
