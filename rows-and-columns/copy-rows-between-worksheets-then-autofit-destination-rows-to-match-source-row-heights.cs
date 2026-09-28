// Title: Copy rows from one worksheet to another and auto‑fit row heights using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that copies a range of rows from a source worksheet to a target worksheet with Aspose.Cells and then applies AutoFitRows to the destination rows. | Demonstrate how to preserve custom row heights when transferring rows between two workbooks using Aspose.Cells and ensure the rows are automatically adjusted after the copy.
// Common Searches: Aspose.Cells C# copy rows from one worksheet to another preserving row height | How to auto‑fit rows after copying them to a different workbook with Aspose.Cells | Copy all rows between workbooks and adjust row heights automatically using Aspose.Cells for .NET
// Tags: Cells.CopyRows method Aspose.Cells | AutoFitRows after row transfer Aspose.Cells | preserve custom row heights C# Aspose.Cells | copy rows between workbooks .NET | row height synchronization Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyAndAutoFit
{
    // The example creates a source workbook with custom row heights, copies all rows to a new workbook using Cells.CopyRows, calls AutoFitRows on the copied range to ensure proper row height adjustment, and saves the result as RowCopyAutoFitResult.xlsx.
    public class Program
    {
        public static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Populate source sheet with sample data
            sourceSheet.Cells["A1"].PutValue("First row data");
            sourceSheet.Cells["A2"].PutValue("Second row data");
            sourceSheet.Cells["A3"].PutValue("Third row data");

            // Set custom row heights in the source sheet
            sourceSheet.Cells.Rows[0].Height = 30; // Row 0 height
            sourceSheet.Cells.Rows[1].Height = 45; // Row 1 height
            sourceSheet.Cells.Rows[2].Height = 20; // Row 2 height

            // ---------- Create destination workbook ----------
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Copy all rows from source to destination (starting at row 0)
            int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount;
            destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, rowsToCopy);

            // Auto‑fit the copied rows in the destination sheet to ensure heights match content
            // (Row heights were already copied, but AutoFitRows guarantees proper adjustment)
            destinationSheet.AutoFitRows(0, rowsToCopy - 1);

            // ---------- Save the result ----------
            destinationWorkbook.Save("RowCopyAutoFitResult.xlsx");
        }
    }
}
