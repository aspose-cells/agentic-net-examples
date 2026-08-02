// Title: Copy Rich‑Text Formatted Cell Between Worksheets with Aspose.Cells for .NET
// Description: Demonstrates how to load a source workbook, locate a rich‑text cell (e.g., A1), and use the Aspose.Cells `Copy` method to transfer its value, formula, and all formatting—including rich‑text segments—to a target cell (e.g., B2) on another worksheet or workbook, then save the result.
// Keywords: Aspose.Cells copy rich text cell | copy cell formatting between worksheets .NET | preserve rich text Aspose.Cells | transfer formatted cell to new workbook | Aspose.Cells C# copy cell with rich text
// Common Searches: copy rich text cell Aspose.Cells C# | how to preserve cell formatting between worksheets | Aspose.Cells copy cell with rich text to another workbook | transfer formatted Excel cell using Aspose.Cells | copy cell value formula and formatting Aspose
// Developer Intent: Copy a cell that contains rich‑text formatting from one worksheet or workbook to another while keeping its value, formula, and styles intact.
// Use Cases: Reuse a styled header from a template workbook across multiple generated reports. | Duplicate a formatted note cell on several sheets of a new workbook without losing color or font variations. | Build a branding workbook by moving rich‑text logo cells from an existing file into a fresh document.
// AI Prompts: Generate C# code using Aspose.Cells that copies a cell with rich‑text formatting from one worksheet to another, preserving formulas and styles. | Show how to copy multiple rich‑text cells from a source sheet to a destination sheet in a different workbook with Aspose.Cells. | Explain how to copy only the rich‑text portion of a cell while leaving other formatting unchanged using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRichTextCopyDemo
{
    // Demonstrates how to load a source workbook, locate a rich‑text cell (e.g., A1), and use the Aspose.Cells `Copy` method to transfer its value, formula, and all formatting—including rich‑text segments—to a target cell (e.g., B2) on another worksheet or workbook, then save the result.
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains the rich‑text formatted cell
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Access the source worksheet and the cell with rich‑text formatting (e.g., A1)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cell sourceCell = sourceSheet.Cells["A1"];

            // Create a new destination workbook (or load an existing one)
            Workbook destinationWorkbook = new Workbook();

            // Access the destination worksheet (first sheet by default) and the target cell (e.g., B2)
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            destinationSheet.Name = "Destination";
            Cell destinationCell = destinationSheet.Cells["B2"];

            // Copy the source cell to the destination cell.
            // This copies the value, formula, and all formatting including rich‑text formatting.
            destinationCell.Copy(sourceCell);

            // Save the destination workbook with the copied rich‑text formatting
            destinationWorkbook.Save("dest.xlsx");
        }
    }
}
