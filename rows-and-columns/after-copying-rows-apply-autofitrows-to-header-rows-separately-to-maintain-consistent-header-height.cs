// Title: Copy rows and auto‑fit header rows with Aspose.Cells for .NET
// Description: Demonstrates how to copy all rows from a source worksheet to a destination worksheet using `Worksheet.Cells.CopyRows`, then apply `Worksheet.AutoFitRows` only to the header rows (first two rows) to keep their height uniform, with optional auto‑fit for the remaining data rows, and finally save the workbook.
// Keywords: Aspose.Cells copy rows | AutoFitRows header | C# Aspose.Cells copy rows | Worksheet.AutoFitRows range | preserve header height | copy worksheet rows .NET
// Common Searches: Aspose.Cells copy rows between workbooks | AutoFitRows only header rows C# | How to keep header height after copying rows Aspose.Cells | Copy rows and auto‑fit specific rows Aspose.Cells
// Developer Intent: Copy every row from one worksheet to another and auto‑fit only the header rows so their height stays consistent.
// Use Cases: Generate a new report by cloning a template sheet and ensuring the header rows retain a fixed height. | Create a summary workbook that reuses rows from an existing file while applying distinct auto‑fit rules to headers and data. | Migrate data to a fresh workbook and format header rows separately for a clean, professional appearance.
// AI Prompts: Show C# code that copies all rows from one worksheet to another with Aspose.Cells and auto‑fits only the first two rows. | How can I use Worksheet.AutoFitRows to adjust header rows after copying rows between workbooks in .NET? | Provide an Aspose.Cells example that copies rows and applies different AutoFitRows ranges for headers and data.

using System;
using Aspose.Cells;

// Demonstrates how to copy all rows from a source worksheet to a destination worksheet using `Worksheet.Cells.CopyRows`, then apply `Worksheet.AutoFitRows` only to the header rows (first two rows) to keep their height uniform, with optional auto‑fit for the remaining data rows, and finally save the workbook.
class AutoFitHeaderAfterCopy
{
    static void Main()
    {
        // Create a source workbook and populate it with header and data rows
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Cells["A1"].PutValue("Header 1");
        sourceSheet.Cells["B1"].PutValue("Header 2");
        sourceSheet.Cells["A2"].PutValue("Data Row 1 - Col A");
        sourceSheet.Cells["B2"].PutValue("Data Row 1 - Col B");
        sourceSheet.Cells["A3"].PutValue("Data Row 2 - Col A");
        sourceSheet.Cells["B3"].PutValue("Data Row 2 - Col B");

        // Create a destination workbook where rows will be copied to
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Copy all rows from the source sheet to the destination sheet
        // Parameters: source cells, source start row, destination start row, number of rows to copy
        int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount;
        destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, rowsToCopy);

        // AutoFit only the header rows (rows 0 and 1) to keep their height consistent
        destinationSheet.AutoFitRows(0, 1);

        // Optionally, AutoFit the remaining data rows
        if (rowsToCopy > 2)
        {
            destinationSheet.AutoFitRows(2, rowsToCopy - 1);
        }

        // Save the resulting workbook
        destinationWorkbook.Save("CopiedWithHeaderAutoFit.xlsx");
    }
}
