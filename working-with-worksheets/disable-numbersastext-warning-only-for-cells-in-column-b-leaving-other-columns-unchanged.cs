// Title: Disable Numbers‑Stored‑as‑Text warning for column B using Aspose.Cells for .NET
// Description: Demonstrates how to add an ErrorCheckOption to a worksheet, turn off the Numbers stored as text warning for the entire column B (B1:B1048576) while keeping all other error‑check settings intact, and save the workbook.
// Keywords: Aspose.Cells | .NET | C# | ErrorCheckOption | Numbers stored as text | disable warning column B | Excel error check range | suppress NumberStoredAsText | column‑specific error check
// Common Searches: Aspose.Cells disable Numbers stored as text for a single column | C# turn off NumberStoredAsText warning column B Aspose.Cells | set error‑check options for specific range Aspose.Cells | suppress NumbersAsText warning in Excel workbook using Aspose.Cells | how to apply column‑level error checks with Aspose.Cells
// Developer Intent: Turn off the Numbers‑stored‑as‑text warning only for cells in column B, leaving all other columns unchanged.
// Use Cases: A financial report where column B holds account numbers stored as text; suppress the warning to avoid user confusion. | A data‑export template that must hide the Numbers‑as‑text alert for a designated column while preserving default checks elsewhere. | Automated workbook generation that applies column‑specific error‑check settings to meet corporate formatting standards.
// AI Prompts: Generate C# code with Aspose.Cells that disables the Numbers stored as text warning for column C only. | Show how to add multiple ErrorCheckOption entries to disable different warnings for separate column ranges in the same worksheet. | Explain how to read existing ErrorCheckOption ranges and toggle the Numbers stored as text warning based on user input at runtime.

using Aspose.Cells;
using System;

// Demonstrates how to add an ErrorCheckOption to a worksheet, turn off the Numbers stored as text warning for the entire column B (B1:B1048576) while keeping all other error‑check settings intact, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the collection of error‑check options for the worksheet
        ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

        // Add a new ErrorCheckOption to the collection
        int optionIndex = errorCheckOptions.Add();
        ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

        // Disable the “Numbers stored as text” warning for this option
        // (ErrorCheckType.NumberStoredAsText and its alias TextNumber)
        errorCheckOption.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false);
        errorCheckOption.SetErrorCheck(ErrorCheckType.TextNumber, false);

        // Define a range that covers the entire column B (from row 1 to the last possible row)
        CellArea columnBRange = CellArea.CreateCellArea("B1", "B1048576");
        errorCheckOption.AddRange(columnBRange);

        // Save the workbook
        workbook.Save("DisableNumbersAsText.xlsx");
    }
}
