// Title: Create an XLSX workbook, merge cells G10:H12, apply a custom date format, and write it to a MemoryStream using Aspose.Cells for .NET
// Common Searches: how to merge a range of cells and format a date in Aspose.Cells C# | saving an Aspose.Cells workbook to a MemoryStream as XLSX in .NET | apply custom date style to merged cells with Aspose.Cells API | Aspose.Cells example for merging G10:H12 and setting dd-mmm-yyyy format | C# code to put current date into merged cells and export to a stream
// Tags: Aspose.Cells merge range G10:H12 | Aspose.Cells custom date format dd-mmm-yyyy | Aspose.Cells export workbook to MemoryStream | Aspose.Cells set cell style programmatically | Aspose.Cells create XLSX in .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, merges the range G10:H12 on the first worksheet, inserts the current date into the merged cell, applies the custom format "dd-mmm-yyyy", and saves the workbook to a MemoryStream in XLSX format using Aspose.Cells for .NET.
public class AsposeCellsExample
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells G10:H12 (zero‑based indices: row 9, column 6, 3 rows, 2 columns)
        sheet.Cells.Merge(9, 6, 3, 2);

        // Put a date value into the merged cell (top‑left cell of the range)
        sheet.Cells["G10"].PutValue(DateTime.Now);

        // Apply a custom date format to the merged cell
        Style style = sheet.Cells["G10"].GetStyle();
        style.Custom = "dd-mmm-yyyy";   // e.g., 06-Nov-2026
        sheet.Cells["G10"].SetStyle(style);

        // Save the workbook to a memory stream in XLSX format
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            // The stream now contains the workbook data; reset position if needed
            stream.Position = 0;

            // Example: you could return the stream, write it to a file, etc.
            // For demonstration, we'll write the stream to a file (optional)
            // File.WriteAllBytes("MergedDateWorkbook.xlsx", stream.ToArray());
        }
    }
}
