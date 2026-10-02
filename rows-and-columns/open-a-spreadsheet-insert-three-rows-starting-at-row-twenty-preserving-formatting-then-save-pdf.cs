// Title: Insert three rows at row 20 with formatting preserved and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing XLSX file, inserts three rows starting at row 20 while copying the original row formatting, and saves the result as a PDF with Aspose.Cells. | Generate a program that uses Aspose.Cells to add rows with style inheritance at a specific index and then converts the workbook to PDF. | Provide a C# example that demonstrates inserting rows with formatting preservation and exporting the modified worksheet to a PDF file.
// Common Searches: Aspose.Cells C# insert rows at specific row number preserving styles | How to add rows to an Excel file and keep formatting before converting to PDF using Aspose.Cells | C# code to insert multiple rows at row 20 and export workbook as PDF with Aspose.Cells | Preserve cell formatting when inserting rows in Aspose.Cells and save as PDF | Insert rows with formatting and generate PDF from worksheet in .NET
// Tags: insert rows with formatting Aspose.Cells C# | convert worksheet to PDF Aspose.Cells | preserve cell styles during row insertion | Aspose.Cells row insertion at specific index | export Excel to PDF after modifying rows

using Aspose.Cells;
using System;

// // Loads input.xlsx, inserts three rows at row 20 while copying existing formatting, and saves the modified workbook as output.pdf in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can also use the worksheet name)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert three rows starting at row 20 (zero‑based index 19)
        // The third argument 'true' copies the formatting from the existing rows
        sheet.Cells.InsertRows(19, 3, true);

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
