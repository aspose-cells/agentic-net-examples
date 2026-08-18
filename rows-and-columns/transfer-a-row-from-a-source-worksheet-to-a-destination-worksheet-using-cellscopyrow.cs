// Title: Copy a Row Between Worksheets Using Cells.CopyRow in Aspose.Cells for .NET (C#)
// Description: The sample creates a source workbook with header and data rows, then copies the second data row into a new destination workbook using the Cells.CopyRow method and saves both files to demonstrate row transfer.
// Keywords: Aspose.Cells | Cells.CopyRow | C# | copy row between worksheets | transfer row between workbooks | Excel automation .NET | worksheet row copy | Aspose.Cells example | row index copy | Excel file manipulation
// Common Searches: Aspose.Cells copy row from one sheet to another | Cells.CopyRow C# example | how to move a row between workbooks using Aspose.Cells | copy row by index Aspose.Cells .NET | transfer Excel row programmatically C#
// Developer Intent: Move a row from a source worksheet to a destination worksheet with the Cells.CopyRow API.
// Use Cases: Duplicate a formatted template row into a newly generated report sheet. | Extract a user‑selected data row from an input file and append it to a summary workbook. | Synchronize row layouts across multiple worksheets in a multi‑sheet workbook.
// AI Prompts: Generate C# code that copies several consecutive rows from one worksheet to another using Cells.CopyRow. | Show how to copy a row while preserving its styles, formulas, and merged cells with Aspose.Cells. | Explain how to adjust row height after copying a row between workbooks in Aspose.Cells.

using System;
using Aspose.Cells;

// The sample creates a source workbook with header and data rows, then copies the second data row into a new destination workbook using the Cells.CopyRow method and saves both files to demonstrate row transfer.
class TransferRowExample
{
    static void Main()
    {
        // Create a source workbook and add sample data
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Header row
        sourceSheet.Cells["A1"].PutValue("ID");
        sourceSheet.Cells["B1"].PutValue("Name");

        // Data rows
        sourceSheet.Cells["A2"].PutValue(1);
        sourceSheet.Cells["B2"].PutValue("Alice");
        sourceSheet.Cells["A3"].PutValue(2);
        sourceSheet.Cells["B3"].PutValue("Bob");

        // Create a destination workbook
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
        destinationSheet.Name = "Destination";

        // Transfer the second row (index 1) from source to the first row (index 0) of destination
        // Using the Cells.CopyRow method as defined in the Aspose.Cells API
        destinationSheet.Cells.CopyRow(sourceSheet.Cells, 1, 0);

        // Save the workbooks to verify the result
        sourceWorkbook.Save("Source.xlsx");
        destinationWorkbook.Save("Destination.xlsx");
    }
}
