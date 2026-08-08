// Title: Copy rows with images and drawing objects using Aspose.Cells for .NET (C#)
// Description: Shows how to transfer all used rows from a source worksheet to a destination worksheet with Aspose.Cells for .NET, keeping cell data, formatting, embedded pictures, charts, and other drawing objects via the default Worksheet.Cells.CopyRows method.
// Keywords: Aspose.Cells copy rows C# | preserve images when copying rows | copy worksheet rows with drawing objects | Worksheet.Cells.CopyRows default behavior | duplicate rows including pictures Aspose
// Common Searches: Aspose.Cells copy rows with pictures | how to keep images when copying rows .NET | copy rows between workbooks preserving shapes | default CopyRows method Aspose.Cells | C# copy worksheet rows including charts
// Developer Intent: Transfer rows from one worksheet to another while automatically retaining embedded visual elements.
// Use Cases: Generate a report that reuses a template’s rows, logos, and charts without manual image handling. | Create a summary workbook that mirrors detailed data plus its associated graphics. | Archive a worksheet’s content with all visual cues intact for compliance or audit purposes.
// AI Prompts: Write C# code that copies a specific row range between worksheets using Aspose.Cells and ensures all pictures and shapes are kept. | Explain the default handling of embedded drawing objects in Worksheet.Cells.CopyRows and list any optional CopyOptions. | Provide a checklist for troubleshooting missing images after using CopyRows in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyExample
{
    // Shows how to transfer all used rows from a source worksheet to a destination worksheet with Aspose.Cells for .NET, keeping cell data, formatting, embedded pictures, charts, and other drawing objects via the default Worksheet.Cells.CopyRows method.
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains data, images and other drawing objects
            Workbook sourceWorkbook = new Workbook("SourceWorkbook.xlsx");
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // first worksheet

            // Create a new workbook that will receive the copied rows
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0]; // first worksheet

            // Define the range of rows to copy
            int sourceStartRow = 0;               // first row in source (zero‑based)
            int destinationStartRow = 0;          // where to paste in destination
            int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount; // copy all used rows

            // Perform the copy using the default CopyRows method (no CopyOptions)
            // This copies cell data, formats, and also embedded images/drawing objects.
            destinationSheet.Cells.CopyRows(
                sourceSheet.Cells,
                sourceStartRow,
                destinationStartRow,
                rowsToCopy);

            // Save the result
            destinationWorkbook.Save("DestinationWorkbook.xlsx");
        }
    }
}
