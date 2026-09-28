// Title: Insert five rows at row 31 with formatting copied from the above row and export the worksheet to PDF using Aspose.Cells for .NET
// AI Prompts: Insert five new rows starting at row 31 in an existing Excel file, copy the formatting of the preceding rows, and save the workbook as a PDF using Aspose.Cells in C#. | Add multiple rows at a specific index while preserving the original row style and updating formulas, then generate a PDF output with Aspose.Cells for .NET. | Use InsertOptions to copy formatting from the row above when inserting rows, and export the modified worksheet to PDF via Workbook.Save.
// Common Searches: how to insert rows with same formatting in Aspose.Cells C# | Aspose.Cells insert multiple rows at specific position and export to PDF | preserve cell references when adding rows using Aspose.Cells .NET | copy row style when inserting rows in Excel workbook with Aspose.Cells | save modified workbook as PDF after inserting rows Aspose.Cells
// Tags: insert rows with formatting Aspose.Cells | copy format from above row C# | save workbook as PDF Aspose.Cells | update references after row insertion Aspose.Cells | insert multiple rows at specific index .NET

using System;
using Aspose.Cells;

// Loads an existing Excel workbook, inserts five rows at the 31st row copying the formatting from the row above, updates cell references, and saves the result as a PDF.
class InsertRowsAndExportPdf
{
    static void Main()
    {
        // Path to the existing workbook
        string inputFile = "input.xlsx";

        // Load the workbook (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(inputFile);

        // Configure insertion options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert 5 rows starting at row index 30 (31st row, zero‑based indexing)
        // The formatting of the rows above will be copied to the new rows
        workbook.Worksheets[0].Cells.InsertRows(30, 5, insertOptions);

        // Save the modified workbook as PDF (uses Workbook.Save(string, SaveFormat))
        string outputPdf = "output.pdf";
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}
