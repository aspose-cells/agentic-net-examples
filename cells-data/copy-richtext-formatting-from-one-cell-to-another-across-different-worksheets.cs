// Title: Copy Rich‑Text Formatting Between Cells Across Worksheets with Aspose.Cells for .NET
// Description: Shows how to load a source workbook, copy a cell that contains rich‑text (bold, italic, color, etc.) to another worksheet or workbook using the Cell.Copy method, and save the updated file while preserving all formatting.
// Keywords: Aspose.Cells | C# | Cell.Copy | rich text formatting | preserve cell formatting | copy cell between workbooks | Excel automation .NET | transfer rich‑text | worksheet to worksheet copy | copy formatted cell
// Common Searches: Aspose.Cells copy cell with rich text | preserve rich‑text when copying Excel cells C# | copy formatted cell to another worksheet Aspose | transfer rich‑text formatting between workbooks .NET | Cell.Copy rich text example
// Developer Intent: Copy a cell’s value and its rich‑text formatting from a source worksheet to a target cell in a different worksheet or workbook using Aspose.Cells for .NET.
// Use Cases: Clone styled header rows from a template workbook into generated reports without losing bold or colored segments. | Replicate title cells with mixed formatting across multiple sheets when building multi‑page invoices. | Move user‑entered rich‑text comments from a data sheet to a summary sheet while keeping all visual cues.
// AI Prompts: Provide C# code that copies a cell with rich‑text formatting from one workbook to another using Aspose.Cells. | Show how to copy only the rich‑text formatting (not the cell value) between worksheets with Aspose.Cells for .NET. | Explain how Cell.Copy preserves rich‑text and what additional steps are needed for merged cells or conditional formatting.

using System;
using Aspose.Cells;

// Shows how to load a source workbook, copy a cell that contains rich‑text (bold, italic, color, etc.) to another worksheet or workbook using the Cell.Copy method, and save the updated file while preserving all formatting.
class CopyRichTextFormatting
{
    static void Main()
    {
        // Load the source workbook that contains the rich‑text formatted cell.
        Workbook sourceWorkbook = new Workbook("Source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        Cell sourceCell = sourceSheet.Cells["A1"]; // cell with rich‑text

        // Create (or load) the destination workbook.
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
        Cell destinationCell = destinationSheet.Cells["B2"]; // target cell

        // Copy the cell value together with all formatting, including rich‑text.
        destinationCell.Copy(sourceCell);

        // Save the workbook with the copied formatting.
        destinationWorkbook.Save("Result.xlsx");
    }
}
