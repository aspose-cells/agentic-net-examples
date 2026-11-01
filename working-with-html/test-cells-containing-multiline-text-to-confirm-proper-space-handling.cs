// Title: Check that spaces in multi‑line cell text are retained after saving and loading with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that writes a multi‑line string with leading, trailing, and internal spaces into a worksheet cell, enables text wrapping, saves the workbook to a MemoryStream, reloads it, and asserts that the cell value matches the original string exactly. | Create a C# unit‑test that inserts multi‑line text containing varied whitespace into cell A1 with Aspose.Cells, performs a save/load round‑trip via a memory stream, and verifies each line’s spaces are unchanged.
// Common Searches: Aspose.Cells .NET keep whitespace in multiline Excel cell after save | C# verify line breaks and spaces are preserved in Excel cell using Aspose.Cells | unit test Aspose.Cells memory stream round‑trip for cell A1 multiline text | how to assert exact whitespace in cell value with Aspose.Cells C# | text wrapping and space handling in Aspose.Cells workbook
// Tags: multiline text whitespace preservation Aspose.Cells | text wrapping cell A1 Aspose.Cells | save workbook to memory stream Aspose.Cells | load workbook from stream verify cell value Aspose.Cells | debug.assert whitespace check Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Diagnostics;

// The example creates a workbook, inserts a multi‑line string with leading, trailing, and internal spaces into cell A1, enables text wrapping, saves the file to a MemoryStream as XLSX, reloads it, and uses Debug.Assert to confirm that the loaded string matches the original line‑by‑line, ensuring whitespace is preserved throughout the save/load cycle.
class MultiLineTextTest
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // Define multi‑line text containing various spaces
        string multiLine = "Line1 with spaces   \nLine2  with  more spaces\n   Line3 leading spaces";

        // Put the multi‑line text into cell A1
        var cell = sheet.Cells["A1"];
        cell.PutValue(multiLine);

        // Enable text wrapping so the line breaks are retained when displayed
        var style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // Save the workbook to a memory stream (in‑memory test, no file I/O)
        using (var ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0; // Reset stream position for reading

            // Load the workbook back from the memory stream
            var loadedWorkbook = new Workbook(ms);
            var loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];
            string loadedValue = loadedCell.StringValue;

            // Verify that the loaded value matches the original multi‑line text
            Debug.Assert(loadedValue == multiLine, "Multi‑line text does not match after load/save.");

            // Additional checks: ensure spaces are preserved on each line
            var lines = loadedValue.Split('\n');
            Debug.Assert(lines[0] == "Line1 with spaces   ", "First line spaces mismatch.");
            Debug.Assert(lines[1] == "Line2  with  more spaces", "Second line spaces mismatch.");
            Debug.Assert(lines[2] == "   Line3 leading spaces", "Third line spaces mismatch.");

            Console.WriteLine("All multi‑line space handling tests passed.");
        }
    }
}
