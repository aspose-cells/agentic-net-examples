// Title: Test Multi‑Line Cell Whitespace Preservation with Aspose.Cells for .NET
// Description: Creates an Excel workbook, writes a cell containing newline characters and multiple spaces, enables text wrapping, auto‑fits the row, saves the file, reloads it, and verifies that line breaks and extra spaces are retained unchanged.
// Keywords: Aspose.Cells | C# | .NET | multi‑line cell | whitespace preservation | extra spaces | text wrapping | Excel round‑trip test | cell StringValue verification
// Common Searches: Aspose.Cells preserve spaces in Excel cell | verify multi‑line text after saving with Aspose.Cells | check whitespace retention in loaded workbook cell | C# test line breaks and spaces in Excel using Aspose
// Developer Intent: Confirm that multi‑line cell content with intentional extra spaces remains identical after a save‑load cycle using Aspose.Cells for .NET.
// Use Cases: Automated unit test that writes a string with line breaks and multiple spaces to a cell, saves the workbook, reloads it, and asserts exact whitespace equality. | Generating reports where user‑entered multi‑line text must keep its original spacing, with programmatic validation of the output. | Debugging scenarios where Excel rendering appears to collapse spaces, using Aspose.Cells to inspect the raw cell value.
// AI Prompts: Generate an NUnit test in C# that uses Aspose.Cells to assert that a cell containing multi‑line text with extra spaces retains the exact whitespace after saving and loading. | Provide C# code that compares the original multi‑line string with the loaded cell's StringValue and logs any differences in spaces or line breaks. | Suggest Aspose.Cells style settings to ensure leading, trailing, and internal spaces are preserved when exporting to Excel.

using System;
using Aspose.Cells;

namespace AsposeCellsMultiLineSpaceTest
{
    // Creates an Excel workbook, writes a cell containing newline characters and multiple spaces, enables text wrapping, auto‑fits the row, saves the file, reloads it, and verifies that line breaks and extra spaces are retained unchanged.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set multi‑line text with intentional extra spaces
            Cell cell = cells["A1"];
            cell.PutValue("First line   with spaces\nSecond line    more spaces");

            // Enable text wrapping so the line break is respected
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Adjust row height to display wrapped text
            worksheet.AutoFitRow(0);

            // Save the workbook (save rule)
            string filePath = "MultiLineSpaceDemo.xlsx";
            workbook.Save(filePath);

            // Load the workbook back (load rule) to verify the content
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            string loadedText = loadedWorksheet.Cells["A1"].StringValue;

            // Output the loaded text to confirm spaces and line breaks are preserved
            Console.WriteLine("Loaded cell text:");
            Console.WriteLine(loadedText);
        }
    }
}
