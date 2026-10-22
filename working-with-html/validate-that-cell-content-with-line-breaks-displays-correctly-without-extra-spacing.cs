// Title: Validate line break preservation and text wrapping in an Excel cell with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, writes a CRLF‑separated string to a cell, enables text wrapping, saves the workbook to a MemoryStream as XLSX, reloads it, and asserts that the three original lines are retained. | Write a C# example using Aspose.Cells that demonstrates how to confirm multiline cell content displays correctly after saving and loading, including style configuration and Debug.Assert checks for each line.
// Common Searches: aspocells preserve newline characters in cell after saving workbook | c# check text wrapping for multiline cell in Excel using Aspose.Cells | validate that Excel cell line breaks are not adding extra rows Aspose.Cells .NET | how to test line break handling in Aspose.Cells memory stream
// Tags: Aspose.Cells multiline cell validation | C# text wrapping Excel cell | preserve CRLF line breaks Aspose.Cells | memory stream workbook save load Aspose.Cells | Debug.Assert multiline cell content .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, inserts a string with CRLF line breaks into cell A1, enables text wrapping, saves the file to a MemoryStream as XLSX, reloads the workbook, reads the cell value, splits it by line‑break characters, and uses Debug.Assert to verify that the three lines match the original content, ensuring correct display without extra spacing.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define cell content with line breaks
        string cellText = "Line1\r\nLine2\r\nLine3";

        // Set the value of cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(cellText);

        // Enable text wrapping so line breaks are displayed
        Style style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // Save the workbook to a memory stream (no file I/O)
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0; // Reset stream position for reading

            // Load the workbook back from the stream
            Workbook loadedWorkbook = new Workbook(ms);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Cell loadedCell = loadedSheet.Cells["A1"];

            // Retrieve the cell value as a string
            string loadedText = loadedCell.StringValue;

            // Validate that the line breaks are preserved
            string[] lines = loadedText.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            Debug.Assert(lines.Length == 3, "Expected 3 lines after splitting by line breaks.");
            Debug.Assert(lines[0] == "Line1", "First line mismatch.");
            Debug.Assert(lines[1] == "Line2", "Second line mismatch.");
            Debug.Assert(lines[2] == "Line3", "Third line mismatch.");

            // If assertions pass, the cell content displays correctly without extra spacing
            Console.WriteLine("Cell content with line breaks validated successfully.");
        }
    }
}
