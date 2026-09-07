// Title: Verify that multi‑line text with spaces is preserved when saving and loading an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a C# example that writes a multi‑line string containing leading, trailing, and internal spaces to cell A1, saves the workbook as .xlsx, reloads it, and asserts that the loaded string matches the original exactly. | Modify the program to test both "\n" and "\r\n" line separators and confirm that whitespace is retained after each save‑load cycle. | Add error handling that logs a detailed message when the cell value after reload differs from the original multi‑line text.
// Common Searches: Aspose.Cells C# preserve spaces and line breaks when writing multi‑line text to a cell | how to test that Excel cell retains whitespace after saving with Aspose.Cells | compare original and loaded cell string with spaces using Aspose.Cells .NET | verify multi‑line string integrity in Excel workbook after reload Aspose.Cells | C# Aspose.Cells line break handling in cell values
// Tags: cell.PutValue multi‑line Aspose.Cells | Workbook.Save preserve cell content | reload workbook cell string verification | Aspose.Cells newline handling | C# Excel cell whitespace validation

using System;
using Aspose.Cells;

// // This program creates a workbook, writes a multi‑line string with varying spaces into cell A1, saves and reloads the file, then compares the original and loaded text to confirm that spaces and line breaks are preserved.
class MultiLineTextTest
{
    static void Main()
    {
        // Create a new workbook (using the standard creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define multi‑line text with spaces and line breaks
        string multiLine = "Line 1 with spaces   \nLine 2  with  more spaces\n   Line 3 leading spaces";

        // Put the multi‑line text into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(multiLine);

        // Save the workbook (using the standard save rule)
        workbook.Save("MultiLineTest.xlsx");

        // Reload the workbook to ensure the text is persisted correctly
        Workbook loadedWorkbook = new Workbook("MultiLineTest.xlsx");
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
        string loadedText = loadedSheet.Cells["A1"].StringValue;

        // Verify that the loaded text matches the original (including spaces and line breaks)
        bool isEqual = string.Equals(multiLine, loadedText, StringComparison.Ordinal);

        // Output the result
        Console.WriteLine("Original Text:");
        Console.WriteLine(multiLine);
        Console.WriteLine("\nLoaded Text:");
        Console.WriteLine(loadedText);
        Console.WriteLine($"\nSpace and line‑break handling correct: {isEqual}");
    }
}
