// Title: Preserve and Verify Cell Line Breaks with Aspose.Cells for .NET
// Description: Shows how to place multiline content into a worksheet cell, turn on text wrapping, auto‑fit the row, save the file, reload it, and ensure the original line‑feed characters stay unchanged and free of extra gaps.
// Keywords: Aspose.Cells C# line break preservation | multiline cell validation .NET | Excel text wrapping after save | compare original and loaded cell value | line feed handling in Excel workbooks | unit test Aspose.Cells line breaks | global Excel automation
// Common Searches: keep line breaks in Excel cells using Aspose.Cells | verify multiline text after workbook reload .NET | Aspose.Cells wrap text and auto‑fit row height | check cell string equality after saving Excel file | C# example for line‑feed retention in cells
// Developer Intent: Ensure that a cell’s multiline string, including its line‑feed markers, is identical before and after the workbook is written to disk and read back.
// Use Cases: Automated regression test confirming that address blocks or comments retain their formatting in generated reports. | Data migration validation where source and destination Excel files must match exactly, line breaks included. | Quality‑control script for templates that rely on wrapped text to appear correctly in final workbooks.
// AI Prompts: Generate a C# unit test with Aspose.Cells that asserts line‑feed preservation in a cell after saving and loading. | Provide code to compare a cell’s StringValue to the original multiline string while handling CRLF vs LF differences. | Explain the steps to enable text wrapping and auto‑fit row height for multiline cells using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsLineBreakValidation
{
    // Shows how to place multiline content into a worksheet cell, turn on text wrapping, auto‑fit the row, save the file, reload it, and ensure the original line‑feed characters stay unchanged and free of extra gaps.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define multi‑line text using line‑feed characters
            string originalText = "First line\nSecond line\nThird line";

            // Put the text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(originalText);

            // Enable text wrapping for the cell so line breaks are respected
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Adjust the row height to show all wrapped lines
            sheet.AutoFitRow(0);

            // Save the workbook to a temporary file
            string filePath = "LineBreakValidation.xlsx";
            workbook.Save(filePath);

            // Reload the workbook from the saved file
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Cell loadedCell = loadedSheet.Cells["A1"];

            // Retrieve the text after loading
            string loadedText = loadedCell.StringValue;

            // Validate that the loaded text matches the original (including line breaks)
            bool isValid = originalText == loadedText;

            Console.WriteLine("Original text:");
            Console.WriteLine(originalText);
            Console.WriteLine("\nLoaded text:");
            Console.WriteLine(loadedText);
            Console.WriteLine($"\nValidation result: {(isValid ? "PASS" : "FAIL")}");
        }
    }
}
