// Title: C# – Disable ‘Numbers Stored as Text’ warning for column B with Aspose.Cells
// Description: Creates a workbook, adds an ErrorCheckOption, turns off the TextNumber (Numbers stored as text) check for the entire column B (B1:B1048576) and saves the file, leaving all other columns unchanged.
// Keywords: Aspose.Cells C# disable Numbers stored as text warning | ErrorCheckOption TextNumber false | disable NumbersAsText column B | Aspose.Cells error check range | C# workbook suppress text-number error | Aspose.Cells column specific error check
// Common Searches: Aspose.Cells turn off Numbers stored as text for one column | C# disable TextNumber error check column B | How to suppress NumbersAsText warning in Aspose.Cells | Set error check options for specific range Aspose.Cells | Aspose.Cells error check per column
// Developer Intent: Turn off the ‘Numbers stored as text’ validation only for cells in column B while preserving default error checks for the rest of the worksheet.
// Use Cases: Hide false ‘Numbers stored as text’ alerts when column B contains identifiers imported as strings, keeping other columns validated normally. | Apply a custom error‑check setting after loading data from an external system that formats numeric values as text in column B. | Configure different error‑check rules per column by adding separate ErrorCheckOption entries for each target range.
// AI Prompts: Show C# code that disables the Numbers stored as text warning for column B using Aspose.Cells, without affecting other columns. | Explain how to use ErrorCheckOptionCollection to apply a TextNumber=false setting to a specific column range in Aspose.Cells. | Give examples of managing multiple ErrorCheckOption entries to set different error‑check behaviors for separate columns in a workbook.

using Aspose.Cells;
using System;

// Creates a workbook, adds an ErrorCheckOption, turns off the TextNumber (Numbers stored as text) check for the entire column B (B1:B1048576) and saves the file, leaving all other columns unchanged.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the collection of error‑check options for the worksheet
        ErrorCheckOptionCollection errorOptions = sheet.ErrorCheckOptions;

        // Add a new option entry
        int optionIndex = errorOptions.Add();
        ErrorCheckOption option = errorOptions[optionIndex];

        // Disable the "Numbers stored as text" warning for this option
        option.SetErrorCheck(ErrorCheckType.TextNumber, false);

        // Define a range that covers the whole column B (from row 1 to the maximum row)
        CellArea columnB = CellArea.CreateCellArea("B1", "B1048576");
        option.AddRange(columnB);

        // Save the workbook (lifecycle: save)
        workbook.Save("NumbersAsTextDisabled.xlsx");
    }
}
