// Title: Copy rich‑text formatting from a source cell to a target cell in another workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy a cell that contains mixed‑font rich text from one workbook to a different workbook while keeping all character styles. | Transfer the styled text segments of cell A1 in a source workbook to cell B2 in a destination workbook using the Cell.Copy method. | Duplicate a cell with multiple character formats into a new Excel file and preserve the original formatting using Aspose.Cells.
// Common Searches: aspnet copy cell rich text to another workbook Aspose.Cells | preserve character‑level formatting when moving Excel cells C# | how to copy a cell with mixed fonts between worksheets using Aspose.Cells | Aspose.Cells copy cell formatting across workbooks example | C# copy rich text cell content to different workbook preserving styles
// Tags: copy cell rich text Aspose.Cells | cell characters formatting copy Aspose.Cells | transfer rich text between workbooks Aspose.Cells | preserve mixed font styles Aspose.Cells | Aspose.Cells copy formatting across worksheets

using System;
using System.IO;
using Aspose.Cells;

// // Demonstrates creating a source workbook with rich‑text in cell A1, applying distinct fonts to substrings, copying that cell to B2 of a new workbook, and saving both workbooks while preserving all character‑level formatting.
class CopyRichTextDemo
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook and add rich‑text to a cell ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cell sourceCell = sourceSheet.Cells["A1"];

            // Set the cell value (the whole text)
            sourceCell.PutValue("Hello World");

            // Apply different font styles to parts of the text (rich‑text)
            // "Hello" part (characters 0‑4)
            sourceCell.Characters(0, 5).Font.Name = "Arial";
            sourceCell.Characters(0, 5).Font.Size = 12;
            sourceCell.Characters(0, 5).Font.IsBold = true;

            // " World" part (characters 5‑10)
            sourceCell.Characters(5, 6).Font.Name = "Times New Roman";
            sourceCell.Characters(5, 6).Font.Size = 14;
            sourceCell.Characters(5, 6).Font.IsItalic = true;

            // ---------- Create destination workbook ----------
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            Cell destinationCell = destinationSheet.Cells["B2"];

            // ---------- Copy the cell (including rich‑text formatting) ----------
            destinationCell.Copy(sourceCell);

            // ---------- Save the workbooks ----------
            string sourcePath = "Source.xlsx";
            string destPath = "Destination.xlsx";

            // Ensure any existing files are overwritten safely
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(destPath)) File.Delete(destPath);

            sourceWorkbook.Save(sourcePath);
            destinationWorkbook.Save(destPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
