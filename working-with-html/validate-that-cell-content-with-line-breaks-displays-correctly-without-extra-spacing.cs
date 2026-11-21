// Title: How to verify that line breaks are retained and displayed without extra spacing in an Excel cell using Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells preserve line breaks in a cell | aspocells verify newline characters after saving workbook | how to enable text wrap for multiline cell using aspocells | check that line breaks are not adding extra spacing in excel with aspocells | validate cell content equality after round‑trip in aspocells c#
// Tags: preserve newline characters in excel cell aspocells | enable text wrapping for multiline cell aspocells | memory stream save load xlsx aspocells | compare cell string value after workbook roundtrip aspocells | validate cell content without extra spacing aspocells

using Aspose.Cells;
using System;

// The example creates a workbook, writes a string containing newline characters to cell A1, enables text wrapping so the line breaks render without additional spacing, saves the workbook to a memory stream as XLSX, reloads it, and compares the reloaded cell value to the original to confirm that line breaks are preserved correctly.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set cell A1 with line breaks
        string cellContent = "Line1\nLine2\nLine3";
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue(cellContent);

        // Enable text wrapping so line breaks are displayed without extra spacing
        Style style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // Save the workbook to a memory stream (no physical file needed)
        using (var stream = new System.IO.MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0;

            // Load the workbook from the memory stream
            Workbook loadedWorkbook = new Workbook(stream);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            string loadedContent = loadedWorksheet.Cells["A1"].StringValue;

            // Validate that the loaded content matches the original with line breaks
            if (loadedContent == cellContent)
            {
                Console.WriteLine("Validation passed: line breaks are preserved correctly.");
            }
            else
            {
                Console.WriteLine("Validation failed: content mismatch.");
                Console.WriteLine($"Expected: [{cellContent}]");
                Console.WriteLine($"Actual:   [{loadedContent}]");
            }
        }
    }
}
