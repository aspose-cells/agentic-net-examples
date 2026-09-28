// Title: Check that a cell with CRLF line breaks retains exact content and wrapping after saving to XLSX in Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to insert a string containing \r\n line breaks into a worksheet cell, enable text wrapping, save the workbook to a MemoryStream as XLSX, reload it, and assert that the cell value matches the original string line by line. | Create a verification routine that loads the saved workbook, reads the cell's StringValue, splits it by CRLF, compares each line to the original array, and confirms that the cell style still has IsTextWrapped set to true.
// Common Searches: how to preserve CRLF line breaks in Aspose.Cells when saving to XLSX | Aspose.Cells verify text wrapping remains after workbook reload | C# check cell content for extra spaces after saving with Aspose.Cells | validate multi-line cell values using memory stream in Aspose.Cells .NET
// Tags: Aspose.Cells line break preservation | C# verify text wrapping after XLSX save | memory stream workbook validation Aspose.Cells | cell content spacing check C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Diagnostics;

// The example creates a workbook, writes a multi-line string with CRLF line breaks into cell A1, enables text wrapping, saves the file to a MemoryStream as XLSX, reloads the workbook, and uses assertions to ensure the original line breaks and wrapping are unchanged and no extra spaces were introduced.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define cell content with line breaks
        string cellContent = "First line\r\nSecond line\r\nThird line";

        // Set the value of cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(cellContent);

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

            // Validate that the cell value still contains the line breaks
            string loadedValue = loadedCell.StringValue;
            Debug.Assert(loadedValue == cellContent, "Cell content mismatch after load.");

            // Validate that there are no extra spaces introduced
            // (Trim each line and compare to original lines)
            string[] originalLines = cellContent.Split(new[] { "\r\n" }, StringSplitOptions.None);
            string[] loadedLines = loadedValue.Split(new[] { "\r\n" }, StringSplitOptions.None);
            Debug.Assert(originalLines.Length == loadedLines.Length, "Line count mismatch.");

            for (int i = 0; i < originalLines.Length; i++)
            {
                Debug.Assert(originalLines[i] == loadedLines[i], $"Line {i + 1} differs after load.");
            }

            // Validate that text wrapping is still enabled
            bool isWrapped = loadedCell.GetStyle().IsTextWrapped;
            Debug.Assert(isWrapped, "Text wrapping is not enabled after load.");

            Console.WriteLine("Validation passed: cell content with line breaks displays correctly without extra spacing.");
        }
    }
}
