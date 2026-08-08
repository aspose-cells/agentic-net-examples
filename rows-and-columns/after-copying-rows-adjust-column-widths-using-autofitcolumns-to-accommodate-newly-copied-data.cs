// Title: Copy rows and AutoFit column widths using Aspose.Cells for .NET (C#)
// Description: Shows how to copy every row from a source worksheet to a destination worksheet with Aspose.Cells, then automatically resize the columns to fit the transferred data via AutoFitColumns, and finally save the workbook as an XLSX file.
// Keywords: Aspose.Cells | CopyRows | AutoFitColumns | C# | .NET | Excel column auto‑fit | copy rows between workbooks | adjust column width | spreadsheet automation | Excel file generation
// Common Searches: Aspose.Cells copy rows and auto fit columns C# | AutoFitColumns after CopyRows .NET | how to copy all rows to another workbook Aspose.Cells | resize columns after copying rows Aspose.Cells | C# copy worksheet rows and auto size columns
// Developer Intent: Copy all rows from a source worksheet to a destination worksheet and automatically adjust the destination columns so the copied content fits neatly.
// Use Cases: Create a report by duplicating template rows and ensuring columns are sized for readability. | Migrate data from a master workbook to a new file while preserving layout and applying auto‑fit for clean presentation. | Consolidate rows from multiple source sheets into one sheet and automatically fit each column for consistent formatting.
// AI Prompts: Generate C# code with Aspose.Cells that copies rows from one worksheet to another and then calls AutoFitColumns on the target sheet. | Explain the steps required to ensure column widths are correctly auto‑fitted after using CopyRows in Aspose.Cells. | Suggest a scalable approach to copy rows from several source worksheets into a single destination worksheet and auto‑fit columns for each sheet.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyRowsAndAutoFit
{
    // Shows how to copy every row from a source worksheet to a destination worksheet with Aspose.Cells, then automatically resize the columns to fit the transferred data via AutoFitColumns, and finally save the workbook as an XLSX file.
    public class Program
    {
        public static void Main()
        {
            // Create source workbook and populate some rows
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Cells["A1"].PutValue("Header");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["A2"].PutValue("Row 1");
            sourceSheet.Cells["B2"].PutValue(12345);
            sourceSheet.Cells["A3"].PutValue("Row 2");
            sourceSheet.Cells["B3"].PutValue(67890);

            // Create destination workbook where rows will be copied
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Copy all rows from source to destination using the CopyRows method
            // sourceRowIndex = 0 (first row), destinationRowIndex = 0, rowNumber = total rows in source
            int totalRows = sourceSheet.Cells.MaxDisplayRange.RowCount;
            destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, totalRows);

            // Adjust column widths in the destination worksheet to fit the newly copied data
            destinationSheet.AutoFitColumns();

            // Save the result workbook
            destinationWorkbook.Save("CopiedRows_AutoFitColumns.xlsx");
        }
    }
}
